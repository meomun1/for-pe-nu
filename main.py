"""
Main entry point for the Multi-Objective Optimization for Air Compressor Scheduling
"""

import os
from data_manager import DataManager
from optimization_approaches import OptimizationApproaches
from results_manager import ResultsManager

def main():
    """Main function to run the multi-objective optimization"""
    # Get credentials file path - works both locally and in GitHub Actions
    credentials_file = os.getenv('GOOGLE_CREDENTIALS_FILE', 'TestAPI.json')
    sheets_id = "13NryaKgZyiU0I0dV9rVWhWblwRPhRLzMCkFtPJ9_VMo"
    
    print("Starting multi-objective optimization for air compressor scheduling...")
    print(f"Using Google Sheet ID: {sheets_id}")
    print(f"Using credentials file: {credentials_file}")
    
    # Initialize components
    data_manager = DataManager(sheets_id, credentials_file)
    data_manager.load_data()
    
    # Initialize optimization approaches with the data manager
    optimizer = OptimizationApproaches(data_manager)
    
    # Initialize results manager
    results_manager = ResultsManager(data_manager)
    
    # Run the optimization approaches
    print("\nRunning optimization approaches...")
    ec_first_results = optimizer.solve_preemptive_EC_first()
    pl_first_results = optimizer.solve_preemptive_PL_first()
    ws_results = optimizer.solve_weighted_sum(w_EC=0.7, w_PL=0.3)  # More emphasis on electricity cost
    cp_results = optimizer.solve_compromise_programming(w_EC=0.5, w_PL=0.5)  # Equal weights
    
    # Collect all results
    all_results = {
        'PR_EC_first': ec_first_results,
        'PR_PL_first': pl_first_results,
        'WS': ws_results,
        'CP': cp_results
    }
    
    # Compare and visualize results
    comparison_df = results_manager.compare_approaches(all_results)
    results_manager.plot_comparison(all_results)
    
    # Save results to Google Sheets
    results_manager.save_results_to_sheets(all_results)
    
    # Add Run schedule formatter after optimization
    print("\nFormatting schedule for better readability...")
    from schedule_formatter import ScheduleFormatter
    formatter = ScheduleFormatter(sheets_id, credentials_file)
    formatter.load_data()
    periods_df = formatter.format_schedule()
    formatter.save_formatted_schedule(periods_df)
    formatter.create_daily_schedule_view(periods_df)
    
    print("\nMulti-objective optimization completed successfully!")
    print("The results have been saved to Google Sheets and local plots.")
    print("\nSummary of results:")
    print(comparison_df)


if __name__ == "__main__":
    main()