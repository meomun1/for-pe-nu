"""
Data Manager module for handling Google Sheets data loading and saving
"""

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime


class DataManager:
    def __init__(self, sheet_id, credentials_file):
        """
        Initialize the data manager with Google Sheets credentials
        
        Args:
            sheet_id: The Google Sheet ID
            credentials_file: Path to the Google API credentials JSON file
        """
        self.sheet_id = sheet_id
        self.credentials_file = credentials_file
        self.connect_to_sheets()
        
    def connect_to_sheets(self):
        """Establish connection to Google Sheets"""
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(self.credentials_file, scopes=scope)
        self.client = gspread.authorize(creds)
        self.sheet = self.client.open_by_key(self.sheet_id)
        print(f"Connected to Google Sheet: {self.sheet.title}")
    
    def load_data(self):
        """Load all necessary data from Google Sheets"""
        print("Loading data from Google Sheets...")
        
        # Load machine parameters
        machine_sheet = self.sheet.worksheet("Machines")
        self.machines_df = pd.DataFrame(machine_sheet.get_all_records())
        print(f"Loaded {len(self.machines_df)} machines")
        
        # Load time-of-use electricity prices
        tou_sheet = self.sheet.worksheet("ToUPrices")
        self.tou_df = pd.DataFrame(tou_sheet.get_all_records())
        print(f"Loaded {len(self.tou_df)} time slots with ToU prices")
        
        # Load incentives if any
        incentive_sheet = self.sheet.worksheet("Incentives")
        self.incentive_df = pd.DataFrame(incentive_sheet.get_all_records())
        print(f"Loaded {len(self.incentive_df)} time slots with incentives")
        
        # Load system parameters
        system_sheet = self.sheet.worksheet("SystemParams")
        system_params_records = system_sheet.get_all_records()
        self.system_params = {row['Parameter']: row['Value'] for row in system_params_records}
        print(f"Loaded system parameters: {list(self.system_params.keys())}")
        
        # Parse parameters
        self.parse_parameters()
    
    def parse_parameters(self):
        """Parse loaded data into model parameters"""
        print("Parsing parameters...")
        
        # Time horizon
        self.T = len(self.tou_df)
        
        # Machine parameters
        self.I = len(self.machines_df)
        self.S = len(self.machines_df['SystemID'].unique())
        
        print(f"Time slots: {self.T}, Machines: {self.I}, Systems: {self.S}")
        
        # Create parameter dictionaries
        self.R = {(row['MachineID'], row['SystemID']): float(row['RatedPower']) 
                 for _, row in self.machines_df.iterrows()}
        
        self.N = {(row['MachineID'], row['SystemID']): int(row['OperationSlots']) 
                 for _, row in self.machines_df.iterrows()}
        
        self.E = {(row['MachineID'], row['SystemID']): int(row['EarlyTimeSlot']) 
                 for _, row in self.machines_df.iterrows()}
        
        self.L = {(row['MachineID'], row['SystemID']): int(row['LateTimeSlot']) 
                 for _, row in self.machines_df.iterrows()}
        
        # Time-of-use prices
        self.c = {row['TimeSlot']: float(row['Price']) for _, row in self.tou_df.iterrows()}
        
        # Incentives
        self.o = {row['TimeSlot']: float(row['Incentive']) for _, row in self.incentive_df.iterrows()}
        
        # System parameters
        self.alpha = float(self.system_params.get('alpha', 0.25))  # Default 15 min (0.25 hour) time slots
        self.A = {}
        for s in range(1, self.S + 1):
            budget_key = f'A_{s}'
            if budget_key in self.system_params:
                self.A[s] = float(self.system_params[budget_key])
            else:
                # Default budget if not specified
                self.A[s] = 5000.0
                print(f"Warning: Budget for system {s} not found. Using default value of {self.A[s]}")
    
    def get_machine_dependencies(self):
        """Extract machine dependencies from the data"""
        machine_dependencies = {}
        for _, row in self.machines_df.iterrows():
            if 'PredecessorMachine' in row and not pd.isna(row['PredecessorMachine']):
                machine_i = int(row['MachineID'])
                system_s = int(row['SystemID'])
                predecessor_i = int(row['PredecessorMachine'])
                machine_dependencies[(machine_i, system_s)] = predecessor_i
        
        return machine_dependencies
    
    def save_worksheet_data(self, title, data, rows=20, cols=10, header_text=None):
        """
        Save data to a worksheet, creating it if it doesn't exist
        
        Args:
            title: Worksheet title
            data: List of lists containing the data
            rows: Number of rows for a new worksheet
            cols: Number of columns for a new worksheet
            header_text: Optional header text to add at the top
        """
        try:
            try:
                worksheet = self.sheet.worksheet(title)
                worksheet.clear()
            except gspread.exceptions.WorksheetNotFound:
                worksheet = self.sheet.add_worksheet(title=title, rows=rows, cols=cols)
            
            # Add header if provided
            if header_text:
                worksheet.update_cell(1, 1, header_text)
                worksheet.update(data, range_name='A3')
            else:
                worksheet.update(data)
                
            return True
        except Exception as e:
            print(f"Error saving data to worksheet '{title}': {e}")
            return False