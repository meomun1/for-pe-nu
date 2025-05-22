"""
Utilities module containing helper functions for the optimization process
"""

import os
import time
import json


def ensure_directory_exists(directory_path):
    """
    Ensure that a directory exists, creating it if necessary
    
    Args:
        directory_path: Path to the directory
    """
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)


def timing_decorator(func):
    """
    Decorator to measure execution time of a function
    
    Args:
        func: Function to time
        
    Returns:
        Wrapped function that prints execution time
    """
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        print(f"Function {func.__name__} execution time: {end_time - start_time:.2f} seconds")
        return result
    return wrapper


def save_json(data, file_path):
    """
    Save data to a JSON file
    
    Args:
        data: Data to save
        file_path: Path to save the file
    """
    with open(file_path, 'w') as f:
        json.dump(data, f, indent=4)


def load_json(file_path):
    """
    Load data from a JSON file
    
    Args:
        file_path: Path to the file
        
    Returns:
        Loaded data or None if file doesn't exist
    """
    if not os.path.exists(file_path):
        return None
    
    with open(file_path, 'r') as f:
        return json.load(f)


def format_results_for_display(results):
    """
    Format optimization results for display
    
    Args:
        results: Dictionary with optimization results
        
    Returns:
        Formatted string representation
    """
    if not results:
        return "No results available"
    
    output = []
    output.append("Optimization Results:")
    output.append(f"- Electricity Cost: {results['EC']:.2f}")
    output.append(f"- Peak Load: {results['PL']:.2f}")
    
    # Count number of machines that are ON
    if 'Schedule' in results:
        machines_on = results['Schedule'][results['Schedule']['Status'] == 1]
        unique_machines = machines_on[['SystemID', 'MachineID']].drop_duplicates()
        output.append(f"- Active Machines: {len(unique_machines)}")
    
    # Add load profile summary
    if 'LoadProfile' in results:
        load_values = list(results['LoadProfile'].values())
        output.append(f"- Average Load: {sum(load_values) / len(load_values):.2f} kW")
        output.append(f"- Maximum Load: {max(load_values):.2f} kW")
        output.append(f"- Minimum Load: {min(load_values):.2f} kW")
    
    return "\n".join(output)


def convert_time_slot_to_time(time_slot, slot_duration_minutes=15):
    """
    Convert a time slot number to a time string (HH:MM)
    
    Args:
        time_slot: Time slot number (1-based)
        slot_duration_minutes: Duration of each time slot in minutes
        
    Returns:
        Time string in HH:MM format
    """
    # Convert inputs to integers to avoid floating point issues
    time_slot = int(time_slot)
    slot_duration_minutes = int(slot_duration_minutes)
    
    # Adjust for 1-based time slots
    minutes_since_midnight = (time_slot - 1) * slot_duration_minutes
    hours = minutes_since_midnight // 60
    minutes = minutes_since_midnight % 60
    
    # Handle day overflow
    hours = hours % 24
    
    return f"{int(hours):02d}:{int(minutes):02d}"