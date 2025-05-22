"""
Results Manager module for extracting, visualizing, and saving optimization results
"""

import pandas as pd
import matplotlib.pyplot as plt
import os
import datetime
import pulp as pl


class ResultExtractor:
    def __init__(self, data_manager):
        """
        Initialize the result extractor
        
        Args:
            data_manager: DataManager instance containing optimization parameters
        """
        self.data_manager = data_manager
    
    def extract_results(self, model, x, y, e, PL, approach_name):
        """
        Extract the objective values and schedule from a solved model
        
        Args:
            model: Solved PuLP model
            x, y: Decision variables
            e: Energy consumption variables
            PL: Peak load variable
            approach_name: Name of the optimization approach
            
        Returns:
            Dictionary with optimization results or None if no optimal solution
        """
        if model.status != pl.LpStatusOptimal:
            print(f"No optimal solution found for {approach_name}")
            return None
        
        # Calculate the electricity cost
        EC = self._calculate_EC_value(x, y)
        
        # Peak load
        peak_load = pl.value(PL)
        
        # Extract machine schedules
        schedule = self._extract_schedule(x)
        
        # Calculate load profile
        load_profile = self._calculate_load_profile(e)
        
        return {
            'EC': EC,
            'PL': peak_load,
            'Schedule': pd.DataFrame(schedule),
            'LoadProfile': load_profile
        }
    
    def _calculate_EC_value(self, x, y):
        """Calculate the actual electricity cost value"""
        R = self.data_manager.R
        c = self.data_manager.c
        o = self.data_manager.o
        alpha = self.data_manager.alpha
        S = self.data_manager.S
        T = self.data_manager.T
        I = self.data_manager.I
        
        cost = 0
        for s in range(1, S + 1):
            for i in range(1, I + 1):
                for t in range(1, T + 1):
                    if (i, s) in R:
                        x_val = pl.value(x[(i, t, s)])
                        y_val = pl.value(y[(i, t, s)])
                        if x_val is not None and y_val is not None:
                            cost += R.get((i, s), 0) * (c.get(t, 0) * x_val - o.get(t, 0) * y_val) * alpha
        
        return cost
    
    def _extract_schedule(self, x):
        """Extract the schedule from decision variables"""
        schedule = []
        S = self.data_manager.S
        I = self.data_manager.I
        T = self.data_manager.T
        R = self.data_manager.R
        
        for s in range(1, S + 1):
            for i in range(1, I + 1):
                for t in range(1, T + 1):
                    if (i, s) in R:  # Only include valid machine-system pairs
                        status = pl.value(x[(i, t, s)])
                        if status is not None and status > 0.5:  # Binary variable = 1
                            schedule.append({
                                'SystemID': s,
                                'MachineID': i,
                                'TimeSlot': t,
                                'Status': 1,
                                'Power': R.get((i, s), 0)
                            })
                        else:
                            schedule.append({
                                'SystemID': s,
                                'MachineID': i,
                                'TimeSlot': t,
                                'Status': 0,
                                'Power': 0
                            })
        
        return schedule
    
    def _calculate_load_profile(self, e):
        """Calculate the load profile from energy variables"""
        T = self.data_manager.T
        load_profile = {}
        
        for t in range(1, T + 1):
            load_profile[t] = pl.value(e[t])
        
        return load_profile


class ResultsManager:
    def __init__(self, data_manager):
        """
        Initialize the results manager
        
        Args:
            data_manager: DataManager instance containing optimization parameters
        """
        self.data_manager = data_manager
    
    def compare_approaches(self, results_dict):
        """
        Compare the results of the different approaches
        
        Args:
            results_dict: Dictionary containing results from different approaches
            
        Returns:
            DataFrame with comparison results
        """
        print("\n=== Comparison of Approaches ===")
        
        if not all(results_dict.values()):
            print("Not all approaches have been solved successfully.")
            return None
        
        # Create comparison table
        comparison = {
            'Approach': [],
            'EC': [],
            'PL': []
        }
        
        for approach, results in results_dict.items():
            comparison['Approach'].append(approach)
            comparison['EC'].append(results['EC'])
            comparison['PL'].append(results['PL'])
        
        df = pd.DataFrame(comparison)
        print(df)
        
        return df
    
    def plot_comparison(self, results_dict):
        """
        Plot the results of the different approaches
        
        Args:
            results_dict: Dictionary containing results from different approaches
        """
        if not all(results_dict.values()):
            print("Not all approaches have been solved successfully.")
            return
        
        # Plot load profiles
        plt.figure(figsize=(14, 8))
        
        for approach, results in results_dict.items():
            load_profile = results['LoadProfile']
            time_slots = list(load_profile.keys())
            loads = list(load_profile.values())
            plt.plot(time_slots, loads, label=f"{approach} (EC={results['EC']:.2f}, PL={results['PL']:.2f})")
        
        plt.xlabel('Time Slot')
        plt.ylabel('Power (kW)')
        plt.title('Load Profiles Comparison')
        plt.legend()
        plt.grid(True)
        
        # Save figure
        os.makedirs('results', exist_ok=True)
        plt.savefig('results/load_profiles_comparison.png')
        plt.close()
        
        # Plot objective values
        approaches = list(results_dict.keys())
        ec_values = [results['EC'] for results in results_dict.values()]
        pl_values = [results['PL'] for results in results_dict.values()]
        
        plt.figure(figsize=(10, 6))
        x = np.arange(len(approaches))
        width = 0.35
        
        plt.bar(x - width/2, ec_values, width, label='Electricity Cost')
        plt.bar(x + width/2, pl_values, width, label='Peak Load')
        
        plt.xlabel('Approach')
        plt.ylabel('Value')
        plt.title('Objective Values Comparison')
        plt.xticks(x, approaches)
        plt.legend()
        plt.grid(True, axis='y')
        
        # Save figure
        plt.savefig('results/objective_values_comparison.png')
        plt.close()
    
    def save_results_to_sheets(self, results_dict):
        """
        Save optimization results to Google Sheets, consolidated into a single sheet
        
        Args:
            results_dict: Dictionary containing results from different approaches
            
        Returns:
            Boolean indicating success or failure
        """
        print("\n=== Saving results to Google Sheets ===")
        
        if not all(results_dict.values()):
            print("Not all approaches have been solved successfully.")
            return False
        
        try:
            # Create comparison table
            comparison = {
                'Approach': [],
                'EC': [],
                'PL': []
            }
            
            for approach, results in results_dict.items():
                comparison['Approach'].append(approach)
                comparison['EC'].append(results['EC'])
                comparison['PL'].append(results['PL'])
            
            comparison_df = pd.DataFrame(comparison)
            
            # Save combined results to a single sheet
            try:
                results_sheet = self.data_manager.sheet.worksheet("OptimizationResults")
                results_sheet.clear()
            except Exception:
                results_sheet = self.data_manager.sheet.add_worksheet(title="OptimizationResults", rows=20, cols=10)
            
            # Add timestamp
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            results_sheet.update_cell(1, 1, f"Multi-Objective Optimization Results - Generated on {timestamp}")
            
            # Add comparison table only (without load profiles)
            results_sheet.update_cell(3, 1, "Comparison of Approaches")
            results_sheet.update('A4', [comparison_df.columns.tolist()] + comparison_df.values.tolist())
            
            # Save the optimized schedule from the selected approach
            # By default, use Weighted Sum (WS) or the first available approach
            selected_approach = 'WS' if 'WS' in results_dict else list(results_dict.keys())[0]
            selected_schedule = results_dict[selected_approach]['Schedule']
            
            # Save to OptimizedSchedule sheet
            try:
                schedule_sheet = self.data_manager.sheet.worksheet("OptimizedSchedule")
                schedule_sheet.clear()
            except Exception:
                schedule_sheet = self.data_manager.sheet.add_worksheet(title="OptimizedSchedule", rows=1000, cols=10)
            
            schedule_sheet.update([selected_schedule.columns.tolist()] + selected_schedule.values.tolist())
            print(f"Updated OptimizedSchedule sheet with {len(selected_schedule)} rows from {selected_approach} approach")
            
            print("Results saved to Google Sheets successfully!")
            return True
        
        except Exception as e:
            print(f"Error saving results to sheets: {e}")
            return False
# Ensure numpy is imported for plotting
import numpy as np