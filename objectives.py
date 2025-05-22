"""
Objectives module for handling objective functions, ideal points, and normalizations
"""

import pulp as pl
from model_builder import ModelBuilder


class ObjectiveManager:
    def __init__(self, data_manager):
        """
        Initialize the objective manager
        
        Args:
            data_manager: DataManager instance containing optimization parameters
        """
        self.data_manager = data_manager
    
    def calculate_EC_expression(self, x, y):
        """
        Calculate the electricity cost expression
        
        Args:
            x: Machine operation decision variables
            y: Incentive application decision variables
            
        Returns:
            Electricity cost expression for the objective function
        """
        R = self.data_manager.R
        c = self.data_manager.c
        o = self.data_manager.o
        alpha = self.data_manager.alpha
        S = self.data_manager.S
        T = self.data_manager.T
        I = self.data_manager.I
        
        return pl.lpSum([R.get((i, s), 0) * 
                        (c.get(t, 0) * x[(i, t, s)] - 
                         o.get(t, 0) * y[(i, t, s)]) * 
                        alpha
                       for s in range(1, S + 1)
                       for t in range(1, T + 1)
                       for i in range(1, I + 1)])
    
    def get_EC_ideal(self):
        """
        Get the ideal value for electricity cost
        
        Returns:
            Minimum possible electricity cost value, or None if optimization fails
        """
        model, x, y, u, e, PL = ModelBuilder.build_base_model(self.data_manager, "EC_Ideal")
        
        # Objective function: Minimize electricity cost
        electricity_cost = self.calculate_EC_expression(x, y)
        model += electricity_cost, "Electricity_Cost"
        
        # Solve the model
        model.solve(pl.PULP_CBC_CMD(msg=False))
        
        if model.status == pl.LpStatusOptimal:
            return pl.value(model.objective)
        else:
            print("Warning: Could not find ideal EC value")
            return None
    
    def get_PL_ideal(self):
        """
        Get the ideal value for peak load
        
        Returns:
            Minimum possible peak load value, or None if optimization fails
        """
        model, x, y, u, e, PL = ModelBuilder.build_base_model(self.data_manager, "PL_Ideal")
        
        # Objective function: Minimize peak load
        model += PL, "Peak_Load"
        
        # Solve the model
        model.solve(pl.PULP_CBC_CMD(msg=False))
        
        if model.status == pl.LpStatusOptimal:
            return pl.value(model.objective)
        else:
            print("Warning: Could not find ideal PL value")
            return None
    
    def get_nadir_points(self):
        """
        Get estimated nadir points for objectives
        
        Returns:
            EC_nadir: Estimated worst electricity cost
            PL_nadir: Estimated worst peak load
        """
        print("Finding ideal points and estimating nadir points...")
        
        # Get ideal values
        EC_ideal = self.get_EC_ideal() or 0
        PL_ideal = self.get_PL_ideal() or 0
        
        # For nadir points, use a more robust approach
        # Run individual objective optimizations to get reasonable ranges
        model_EC, x_EC, y_EC, u_EC, e_EC, PL_EC = ModelBuilder.build_base_model(self.data_manager, "EC_Only")
        model_EC += self.calculate_EC_expression(x_EC, y_EC), "EC_Only"
        model_EC.solve(pl.PULP_CBC_CMD(msg=False))
        
        model_PL, x_PL, y_PL, u_PL, e_PL, PL_PL = ModelBuilder.build_base_model(self.data_manager, "PL_Only")
        model_PL += PL_PL, "PL_Only"
        model_PL.solve(pl.PULP_CBC_CMD(msg=False))
        
        # Get PL value from EC optimization and EC value from PL optimization
        if model_EC.status == pl.LpStatusOptimal and model_PL.status == pl.LpStatusOptimal:
            EC_from_EC_opt = pl.value(self.calculate_EC_expression(x_EC, y_EC))
            PL_from_EC_opt = pl.value(PL_EC)
            
            EC_from_PL_opt = pl.value(self.calculate_EC_expression(x_PL, y_PL))
            PL_from_PL_opt = pl.value(PL_PL)
            
            # Use these as approximations of nadir points
            EC_nadir = max(EC_from_PL_opt, EC_from_EC_opt * 1.5)  # 50% worse than optimal EC
            PL_nadir = max(PL_from_EC_opt, PL_from_PL_opt * 1.5)  # 50% worse than optimal PL
        else:
            # Fallback to safe default values
            EC_nadir = EC_ideal * 2 if EC_ideal else 100  # Double the ideal or default to 100
            PL_nadir = PL_ideal * 2 if PL_ideal else 20   # Double the ideal or default to 20
        
        # Ensure nadir and ideal points are different to avoid division by zero
        if EC_nadir == EC_ideal:
            EC_nadir = EC_ideal * 1.1 + 0.1
        if PL_nadir == PL_ideal:
            PL_nadir = PL_ideal * 1.1 + 0.1
        
        print(f"EC ideal: {EC_ideal:.2f}, EC nadir (estimated): {EC_nadir:.2f}")
        print(f"PL ideal: {PL_ideal:.2f}, PL nadir (estimated): {PL_nadir:.2f}")
        
        return EC_ideal, PL_ideal, EC_nadir, PL_nadir
    
    def get_normalization_factors(self):
        """
        Get normalization factors for objectives
        
        Returns:
            EC_ideal: Ideal electricity cost
            PL_ideal: Ideal peak load
            EC_norm: Normalization factor for electricity cost
            PL_norm: Normalization factor for peak load
        """
        EC_ideal, PL_ideal, EC_nadir, PL_nadir = self.get_nadir_points()
        
        # Calculate normalization factors
        EC_norm = 1 / (EC_nadir - EC_ideal)
        PL_norm = 1 / (PL_nadir - PL_ideal)
        
        return EC_ideal, PL_ideal, EC_norm, PL_norm