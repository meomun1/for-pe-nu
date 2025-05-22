"""
Optimization Approaches module implementing various multi-objective optimization methods
"""

import pulp as pl
from model_builder import ModelBuilder
from objectives import ObjectiveManager
from results_manager import ResultExtractor


class OptimizationApproaches:
    def __init__(self, data_manager):
        """
        Initialize optimization approaches with data manager
        
        Args:
            data_manager: DataManager instance containing optimization parameters
        """
        self.data_manager = data_manager
        self.objective_manager = ObjectiveManager(data_manager)
        self.result_extractor = ResultExtractor(data_manager)
    
    def solve_preemptive_EC_first(self):
        """
        Solve using preemptive approach with EC as first priority
        
        Returns:
            Dictionary with optimization results
        """
        print("\n=== Solving with Preemptive Approach (EC First) ===")
        
        # Step 1: Optimize EC alone
        print("Step 1: Optimizing Electricity Cost (EC)...")
        model1, x1, y1, u1, e1, PL1 = ModelBuilder.build_base_model(self.data_manager, "PR_EC_First_Step1")
        
        # Objective function: Minimize electricity cost
        electricity_cost = self.objective_manager.calculate_EC_expression(x1, y1)
        model1 += electricity_cost, "Electricity_Cost"
        
        # Solve the model
        model1.solve(pl.PULP_CBC_CMD(msg=False))
        
        if model1.status != pl.LpStatusOptimal:
            print("No optimal solution found in Step 1")
            return None
        
        optimal_EC = pl.value(model1.objective)
        print(f"Optimal EC value: {optimal_EC:.2f}")
        
        # Step 2: Optimize PL while maintaining optimal EC
        print("Step 2: Optimizing Peak Load (PL) while maintaining optimal EC...")
        model2, x2, y2, u2, e2, PL2 = ModelBuilder.build_base_model(self.data_manager, "PR_EC_First_Step2")
        
        # Add constraint to maintain optimal EC
        model2 += (self.objective_manager.calculate_EC_expression(x2, y2) <= optimal_EC * 1.001, "Maintain_Optimal_EC")  # Allow 0.1% tolerance
        
        # Objective function: Minimize peak load
        model2 += PL2, "Peak_Load"
        
        # Solve the model
        model2.solve(pl.PULP_CBC_CMD(msg=False))
        
        # Extract results
        results = self.result_extractor.extract_results(model2, x2, y2, e2, PL2, "PR_EC_First")
        if results:
            print(f"PR_EC_First completed. EC: {results['EC']:.2f}, PL: {results['PL']:.2f}")
        
        return results
    
    def solve_preemptive_PL_first(self):
        """
        Solve using preemptive approach with PL as first priority
        
        Returns:
            Dictionary with optimization results
        """
        print("\n=== Solving with Preemptive Approach (PL First) ===")
        
        # Step 1: Optimize PL alone
        print("Step 1: Optimizing Peak Load (PL)...")
        model1, x1, y1, u1, e1, PL1 = ModelBuilder.build_base_model(self.data_manager, "PR_PL_First_Step1")
        
        # Objective function: Minimize peak load
        model1 += PL1, "Peak_Load"
        
        # Solve the model
        model1.solve(pl.PULP_CBC_CMD(msg=False))
        
        if model1.status != pl.LpStatusOptimal:
            print("No optimal solution found in Step 1")
            return None
        
        optimal_PL = pl.value(PL1)
        print(f"Optimal PL value: {optimal_PL:.2f}")
        
        # Step 2: Optimize EC while maintaining optimal PL
        print("Step 2: Optimizing Electricity Cost (EC) while maintaining optimal PL...")
        model2, x2, y2, u2, e2, PL2 = ModelBuilder.build_base_model(self.data_manager, "PR_PL_First_Step2")
        
        # Add constraint to maintain optimal PL
        model2 += (PL2 <= optimal_PL * 1.001, "Maintain_Optimal_PL")  # Allow 0.1% tolerance
        
        # Objective function: Minimize electricity cost
        electricity_cost = self.objective_manager.calculate_EC_expression(x2, y2)
        model2 += electricity_cost, "Electricity_Cost"
        
        # Solve the model
        model2.solve(pl.PULP_CBC_CMD(msg=False))
        
        # Extract results
        results = self.result_extractor.extract_results(model2, x2, y2, e2, PL2, "PR_PL_First")
        if results:
            print(f"PR_PL_First completed. EC: {results['EC']:.2f}, PL: {results['PL']:.2f}")
        
        return results
    
    def solve_weighted_sum(self, w_EC=0.5, w_PL=0.5):
        """
        Solve using weighted sum approach with normalization
        
        Args:
            w_EC: Weight for electricity cost objective (0-1)
            w_PL: Weight for peak load objective (0-1)
            
        Returns:
            Dictionary with optimization results
        """
        print(f"\n=== Solving with Weighted Sum Approach (EC: {w_EC}, PL: {w_PL}) ===")
        
        # Get normalization factors
        EC_ideal, PL_ideal, EC_norm, PL_norm = self.objective_manager.get_normalization_factors()
        
        # Build model
        model, x, y, u, e, PL = ModelBuilder.build_base_model(self.data_manager, "Weighted_Sum")
        
        # Create electricity cost expression
        electricity_cost = self.objective_manager.calculate_EC_expression(x, y)
        
        # Create normalized weighted objective
        # Convert objective expressions to relative deviations from ideal points
        EC_normalized = (electricity_cost - EC_ideal) * EC_norm
        PL_normalized = (PL - PL_ideal) * PL_norm
        
        weighted_objective = w_EC * EC_normalized + w_PL * PL_normalized
        model += weighted_objective, "Weighted_Objective"
        
        # Solve the model
        model.solve(pl.PULP_CBC_CMD(msg=False))
        
        # Extract results
        results = self.result_extractor.extract_results(model, x, y, e, PL, "Weighted_Sum")
        if results:
            print(f"Weighted Sum completed. EC: {results['EC']:.2f}, PL: {results['PL']:.2f}")
        
        return results
    
    def solve_compromise_programming(self, w_EC=0.5, w_PL=0.5):
        """
        Solve using compromise programming approach
        
        Args:
            w_EC: Weight for electricity cost objective (0-1)
            w_PL: Weight for peak load objective (0-1)
            
        Returns:
            Dictionary with optimization results
        """
        print(f"\n=== Solving with Compromise Programming Approach (EC: {w_EC}, PL: {w_PL}) ===")
        
        # Get ideal points with fallback values
        print("Finding ideal points...")
        EC_ideal = self.objective_manager.get_EC_ideal() or 0.1  # Avoid division by zero
        PL_ideal = self.objective_manager.get_PL_ideal() or 0.1  # Avoid division by zero
        
        print(f"EC ideal: {EC_ideal:.2f}, PL ideal: {PL_ideal:.2f}")
        
        # Build model
        model, x, y, u, e, PL = ModelBuilder.build_base_model(self.data_manager, "Compromise_Programming")
        
        # Create electricity cost expression
        electricity_cost = self.objective_manager.calculate_EC_expression(x, y)
        
        # Create variables for deviations from ideal point
        max_dev = pl.LpVariable("Maximum_Deviation", lowBound=0)
        
        # Add deviation constraints - using relative deviations
        # Add small epsilon to avoid division by zero
        model += (w_EC * ((electricity_cost - EC_ideal) / max(EC_ideal, 0.001)) <= max_dev, "EC_Deviation_Constraint")
        model += (w_PL * ((PL - PL_ideal) / max(PL_ideal, 0.001)) <= max_dev, "PL_Deviation_Constraint")
        
        # Set objective to minimize maximum deviation
        model += max_dev, "Minimize_Maximum_Deviation"
        
        # Solve the model
        model.solve(pl.PULP_CBC_CMD(msg=False))
        
        # Extract results
        results = self.result_extractor.extract_results(model, x, y, e, PL, "Compromise_Programming")
        if results:
            print(f"Compromise Programming completed. EC: {results['EC']:.2f}, PL: {results['PL']:.2f}")
        
        return results