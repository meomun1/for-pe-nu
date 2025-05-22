"""
Model Builder module for creating optimization models
"""

import pulp as pl


class ModelBuilder:
    @staticmethod
    def build_base_model(data_manager, name="Base_Model"):
        """
        Build base optimization model with all constraints but no objective
        
        Args:
            data_manager: DataManager instance containing optimization parameters
            name: Name for the optimization model
            
        Returns:
            model: The PuLP optimization model
            x, y, u: Decision variables
            e: Energy consumption variables
            PL: Peak load variable
        """
        # Create the model
        model = pl.LpProblem(name, pl.LpMinimize)
        
        # Extract parameters from data manager
        I = data_manager.I  # Number of machines
        T = data_manager.T  # Number of time slots
        S = data_manager.S  # Number of systems
        R = data_manager.R  # Rated power
        E = data_manager.E  # Early time slot
        L = data_manager.L  # Late time slot
        N = data_manager.N  # Operation slots
        c = data_manager.c  # ToU prices
        o = data_manager.o  # Incentives
        alpha = data_manager.alpha  # Time slot duration
        A = data_manager.A  # Budget
        
        # Create decision variables
        x = pl.LpVariable.dicts("x", 
                               [(i, t, s) for i in range(1, I + 1) 
                                          for t in range(1, T + 1)
                                          for s in range(1, S + 1)],
                               cat=pl.LpBinary)
        
        y = pl.LpVariable.dicts("y", 
                               [(i, t, s) for i in range(1, I + 1) 
                                          for t in range(1, T + 1)
                                          for s in range(1, S + 1)],
                               cat=pl.LpBinary)
        
        u = pl.LpVariable.dicts("u", 
                               [(i, t, s) for i in range(1, I + 1) 
                                          for t in range(1, T + 1)
                                          for s in range(1, S + 1)],
                               cat=pl.LpBinary)
        
        e = pl.LpVariable.dicts("e", range(1, T + 1), lowBound=0)
        PL = pl.LpVariable("PL", lowBound=0)
        
        # Add constraints
        ModelBuilder._add_electricity_consumption_constraints(model, e, x, R, T, S, I)
        ModelBuilder._add_peak_load_constraints(model, PL, e, T)
        ModelBuilder._add_budget_constraints(model, x, y, R, c, o, alpha, A, T, I, S)
        ModelBuilder._add_operation_duration_constraints(model, x, E, L, N, S, I)
        ModelBuilder._add_uninterruptible_operation_constraints(model, x, u, S, I, T)
        ModelBuilder._add_machine_dependency_constraints(model, x, u, data_manager.get_machine_dependencies(), T)
        ModelBuilder._add_incentive_constraints(model, x, y, S, I, T)
        
        return model, x, y, u, e, PL
    
    @staticmethod
    def _add_electricity_consumption_constraints(model, e, x, R, T, S, I):
        """Add electricity consumption calculation constraints"""
        for t in range(1, T + 1):
            model += (e[t] == pl.lpSum([R.get((i, s), 0) * x[(i, t, s)]
                                       for s in range(1, S + 1)
                                       for i in range(1, I + 1)]),
                     f"Electricity_Consumption_{t}")
    
    @staticmethod
    def _add_peak_load_constraints(model, PL, e, T):
        """Add peak load constraints"""
        for t in range(1, T + 1):
            model += (PL >= e[t], f"Peak_Load_{t}")
    
    @staticmethod
    def _add_budget_constraints(model, x, y, R, c, o, alpha, A, T, I, S):
        """Add budget constraints"""
        for s in range(1, S + 1):
            model += (pl.lpSum([R.get((i, s), 0) * 
                              (c.get(t, 0) * x[(i, t, s)] - 
                               o.get(t, 0) * y[(i, t, s)]) * 
                              alpha
                             for t in range(1, T + 1)
                             for i in range(1, I + 1)]) <= A[s],
                     f"Budget_Constraint_{s}")
    
    @staticmethod
    def _add_operation_duration_constraints(model, x, E, L, N, S, I):
        """Add operation duration constraints"""
        for s in range(1, S + 1):
            for i in range(1, I + 1):
                if (i, s) in E and (i, s) in L:
                    model += (pl.lpSum([x[(i, t, s)] 
                                      for t in range(E[(i, s)], L[(i, s)] + 1)]) 
                             >= N.get((i, s), 0),
                            f"Operation_Duration_{i}_{s}")
    
    @staticmethod
    def _add_uninterruptible_operation_constraints(model, x, u, S, I, T):
        """Add uninterruptible operation constraints"""
        for s in range(1, S + 1):
            for i in range(1, I + 1):
                for t in range(1, T + 1):
                    # Constraint (7)
                    model += (x[(i, t, s)] <= 1 - u[(i, t, s)],
                             f"Uninterruptible_1_{i}_{t}_{s}")
                    
                    # Constraint (8) - for t >= 2
                    if t >= 2:
                        model += (x[(i, t-1, s)] - x[(i, t, s)] <= u[(i, t, s)],
                                 f"Uninterruptible_2_{i}_{t}_{s}")
                    
                    # Constraint (9) - for t >= 2
                    if t >= 2:
                        model += (u[(i, t-1, s)] <= u[(i, t, s)],
                                 f"Uninterruptible_3_{i}_{t}_{s}")
    
    @staticmethod
    def _add_machine_dependency_constraints(model, x, u, machine_dependencies, T):
        """Add machine dependency constraints"""
        for (i, s), i_star in machine_dependencies.items():
            for t in range(1, T + 1):
                model += (x[(i, t, s)] <= u[(i_star, t, s)],
                         f"Precedence_Constraint_{i}_{i_star}_{t}_{s}")
    
    @staticmethod
    def _add_incentive_constraints(model, x, y, S, I, T):
        """Add incentive constraints"""
        for s in range(1, S + 1):
            for i in range(1, I + 1):
                for t in range(1, T + 1):
                    model += (y[(i, t, s)] <= x[(i, t, s)], 
                             f"Incentive_Constraint_{i}_{t}_{s}")