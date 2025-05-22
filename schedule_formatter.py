"""
Schedule Formatter for Air Compressor Optimization Results

This script reads the OptimizedSchedule sheet from a Google Sheet and creates
a new human-readable sheet showing when each machine should be turned on and off.
"""

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import numpy as np
from utils import convert_time_slot_to_time


class ScheduleFormatter:
    def __init__(self, sheet_id, credentials_file):
        """
        Initialize the schedule formatter with Google Sheets credentials
        
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
        """Load the required data from Google Sheets"""
        print("Loading data from Google Sheets...")
        
        # Load optimized schedule
        schedule_sheet = self.sheet.worksheet("OptimizedSchedule")
        self.schedule_df = pd.DataFrame(schedule_sheet.get_all_records())
        print(f"Loaded {len(self.schedule_df)} schedule entries")
        
        # Load machine data for names
        machine_sheet = self.sheet.worksheet("Machines")
        self.machines_df = pd.DataFrame(machine_sheet.get_all_records())
        print(f"Loaded {len(self.machines_df)} machine entries")
        
        # Load ToU data for time information
        tou_sheet = self.sheet.worksheet("ToUPrices")
        self.tou_df = pd.DataFrame(tou_sheet.get_all_records())
        
        # Get time slot duration (default 15 minutes)
        system_sheet = self.sheet.worksheet("SystemParams")
        system_params = pd.DataFrame(system_sheet.get_all_records())
        alpha_row = system_params[system_params['Parameter'] == 'alpha']
        self.slot_duration_minutes = 60 * float(alpha_row['Value'].iloc[0]) if not alpha_row.empty else 15
    
    def format_schedule(self):
        """
        Format the optimized schedule into a human-readable format
        
        Returns:
            DataFrame with machine operation periods
        """
        print("Formatting schedule...")
        
        # Create a dictionary to store machine names
        machine_names = {}
        for _, row in self.machines_df.iterrows():
            key = (int(row['SystemID']), int(row['MachineID']))
            name = row.get('MachineName', f"Machine {row['MachineID']}")
            machine_names[key] = name
        
        # Create a list to store machine operation periods
        operation_periods = []
        
        # Identify unique system-machine pairs
        system_machines = self.schedule_df[['SystemID', 'MachineID']].drop_duplicates()
        
        # Process each machine
        for _, sm_row in system_machines.iterrows():
            system_id = int(sm_row['SystemID'])
            machine_id = int(sm_row['MachineID'])
            
            # Filter schedule for this machine
            machine_schedule = self.schedule_df[
                (self.schedule_df['SystemID'] == system_id) & 
                (self.schedule_df['MachineID'] == machine_id)
            ].sort_values('TimeSlot')
            
            # Get machine name
            machine_name = machine_names.get((system_id, machine_id), f"Machine {machine_id}")
            
            # Find start and end times
            status_changes = []
            prev_status = 0
            
            for _, row in machine_schedule.iterrows():
                time_slot = int(row['TimeSlot'])
                status = int(row['Status'])
                
                if status != prev_status:
                    status_changes.append((time_slot, status))
                    prev_status = status
            
            # Process status changes to find operation periods
            for i in range(0, len(status_changes) - 1, 2):
                if i + 1 < len(status_changes):
                    start_slot = status_changes[i][0]
                    end_slot = status_changes[i+1][0] - 1  # The machine runs until the end of the previous slot
                    
                    # Convert slots to time strings
                    start_time = convert_time_slot_to_time(start_slot, self.slot_duration_minutes)
                    end_time = convert_time_slot_to_time(end_slot + 1, self.slot_duration_minutes)  # +1 because we want the end of this slot
                    
                    # Get power consumption
                    power = 0
                    power_row = machine_schedule[machine_schedule['Status'] == 1].iloc[0] if not machine_schedule[machine_schedule['Status'] == 1].empty else None
                    if power_row is not None:
                        power = power_row['Power']
                    
                    # Calculate duration in time slots and minutes
                    duration_slots = end_slot - start_slot + 1
                    duration_minutes = duration_slots * self.slot_duration_minutes
                    
                    # Add to operation periods
                    operation_periods.append({
                        'SystemID': system_id,
                        'MachineID': machine_id,
                        'MachineName': machine_name,
                        'StartSlot': start_slot,
                        'EndSlot': end_slot,
                        'StartTime': start_time,
                        'EndTime': end_time,
                        'DurationSlots': duration_slots,
                        'DurationMinutes': duration_minutes,
                        'Power': power,
                        'EnergyConsumption': power * (duration_slots * self.slot_duration_minutes / 60)  # kWh
                    })
        
        # Convert to DataFrame
        if operation_periods:
            periods_df = pd.DataFrame(operation_periods)
            return periods_df
        else:
            print("No operation periods found. Check if machines are scheduled to run.")
            return pd.DataFrame(columns=[
                'SystemID', 'MachineID', 'MachineName', 'StartSlot', 'EndSlot', 
                'StartTime', 'EndTime', 'DurationSlots', 'DurationMinutes', 
                'Power', 'EnergyConsumption'
            ])
    
    def save_formatted_schedule(self, periods_df):
        """
        Save the formatted schedule to a new worksheet
        
        Args:
            periods_df: DataFrame with machine operation periods
        """
        print("Saving formatted schedule to Google Sheets...")
        
        try:
            # Create a new worksheet or clear existing one
            try:
                schedule_sheet = self.sheet.worksheet("MachineOperationSchedule")
                schedule_sheet.clear()
            except gspread.exceptions.WorksheetNotFound:
                schedule_sheet = self.sheet.add_worksheet(title="MachineOperationSchedule", rows=len(periods_df) + 10, cols=15)
            
            # Add timestamp
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            schedule_sheet.update_cell(1, 1, f"Machine Operation Schedule - Generated on {timestamp}")
            
            # Prepare data
            header = periods_df.columns.tolist()
            data = [header] + periods_df.values.tolist()
            
            # Update sheet
            schedule_sheet.update('A3', data)
            
            # Format the sheet
            try:
                # Set the header row to bold
                bold_format = {
                    "textFormat": {
                        "bold": True
                    }
                }
                
                # Format header row
                header_cells = schedule_sheet.range(3, 1, 3, len(header))
                for cell in header_cells:
                    schedule_sheet.format(f"{cell.address}", bold_format)
                
                print("Formatted the header row")
            except Exception as e:
                print(f"Warning: Could not apply formatting: {str(e)}")
            
            print(f"Saved machine operation schedule to 'MachineOperationSchedule' sheet")
            return True
            
        except Exception as e:
            print(f"Error saving formatted schedule: {str(e)}")
            return False
    
    def create_daily_schedule_view(self, periods_df):
        """
        Create a daily schedule view showing machines and their operation times
        with 10-minute intervals for more precise visualization
        
        Args:
            periods_df: DataFrame with machine operation periods
        """
        print("Creating daily schedule view with 10-minute intervals...")
        
        try:
            # Create a new worksheet or clear existing one
            try:
                daily_sheet = self.sheet.worksheet("DailyScheduleView")
                daily_sheet.clear()
            except gspread.exceptions.WorksheetNotFound:
                daily_sheet = self.sheet.add_worksheet(title="DailyScheduleView", rows=50, cols=150)
            
            # Add timestamp
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            daily_sheet.update_cell(1, 1, f"Daily Machine Schedule (10-min intervals) - Generated on {timestamp}")
            
            # Prepare the daily schedule
            # Header: 10-minute intervals throughout the day
            header = ["Machine/Time"]
            
            # Create header with 10-minute intervals
            for hour in range(24):
                for minute in range(0, 60, 10):
                    header.append(f"{hour:02d}:{minute:02d}")
            
            # Create a row for each machine
            machines = periods_df[['SystemID', 'MachineID', 'MachineName']].drop_duplicates()
            
            # Prepare data rows
            data_rows = []
            for _, machine in machines.iterrows():
                system_id = machine['SystemID']
                machine_id = machine['MachineID']
                machine_name = machine['MachineName']
                
                # Initialize row with machine name
                row = [f"System {system_id} - {machine_name}"]
                
                # Fill with spaces initially (144 intervals of 10 minutes in a day)
                row.extend([""] * 144)
                
                # Get operation periods for this machine
                machine_periods = periods_df[
                    (periods_df['SystemID'] == system_id) & 
                    (periods_df['MachineID'] == machine_id)
                ]
                
                # Mark operation periods
                for _, period in machine_periods.iterrows():
                    start_time = period['StartTime']
                    end_time = period['EndTime']
                    
                    # Extract hours and minutes
                    start_hour = int(start_time.split(':')[0])
                    start_minute = int(start_time.split(':')[1])
                    end_hour = int(end_time.split(':')[0])
                    end_minute = int(end_time.split(':')[1])
                    
                    # Calculate corresponding indices in the 10-minute grid
                    start_index = (start_hour * 6) + (start_minute // 10)
                    
                    # For end time, round up to the next 10-minute slot if needed
                    end_index = (end_hour * 6) + (end_minute // 10)
                    if end_minute % 10 > 0:
                        end_index += 1
                    
                    # Mark operation intervals
                    for interval in range(start_index, end_index):
                        if 0 <= interval < 144:  # Ensure within bounds
                            row[interval + 1] = "ON"  # +1 because first column is machine name
                
                data_rows.append(row)
            
            # Combine header and data rows
            all_data = [header] + data_rows
            
            # Update sheet
            daily_sheet.update('A3', all_data)
            
            # Apply conditional formatting for ON cells
            try:
                formatting_request = {
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{
                                "sheetId": daily_sheet.id,
                                "startRowIndex": 3,
                                "endRowIndex": 3 + len(data_rows),
                                "startColumnIndex": 1,
                                "endColumnIndex": 145  # 144 intervals + 1 for machine name
                            }],
                            "booleanRule": {
                                "condition": {
                                    "type": "TEXT_EQ",
                                    "values": [{"userEnteredValue": "ON"}]
                                },
                                "format": {
                                    "backgroundColor": {"red": 0.7, "green": 0.9, "blue": 0.7}
                                }
                            }
                        },
                        "index": 0
                    }
                }
                
                self.sheet.batch_update({"requests": [formatting_request]})
                print("Applied conditional formatting")
            except Exception as e:
                print(f"Warning: Could not apply conditional formatting: {str(e)}")
            
            print(f"Created daily schedule view in 'DailyScheduleView' sheet")
            return True
            
        except Exception as e:
            print(f"Error creating daily schedule view: {str(e)}")
            return False


def main():
    """Main function to format the schedule"""
    # Replace with your actual file path and Google Sheet ID
    credentials_file = r'/Users/nguyenphuong/Library/CloudStorage/OneDrive-VietNamNationalUniversity-HCMINTERNATIONALUNIVERSITY/Documents/PYTHON/CODE/APPSHEET/TestAPI.json'
    sheets_id = "13NryaKgZyiU0I0dV9rVWhWblwRPhRLzMCkFtPJ9_VMo"
    
    print("Starting schedule formatting...")
    print(f"Using Google Sheet ID: {sheets_id}")
    print(f"Using credentials file: {credentials_file}")
    
    # Create schedule formatter
    formatter = ScheduleFormatter(sheets_id, credentials_file)
    
    # Load data
    formatter.load_data()
    
    # Format schedule
    periods_df = formatter.format_schedule()
    
    # Save formatted schedule
    formatter.save_formatted_schedule(periods_df)
    
    # Create daily schedule view
    formatter.create_daily_schedule_view(periods_df)
    
    print("\nSchedule formatting completed successfully!")
    print("Two new sheets have been created:")
    print("1. MachineOperationSchedule - Detailed operation periods for each machine")
    print("2. DailyScheduleView - Visual daily schedule showing when machines are ON")


if __name__ == "__main__":
    main()