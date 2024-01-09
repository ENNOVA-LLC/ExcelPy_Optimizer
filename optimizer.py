""" Optimizes the parameters in the `pySolve_Param` range.

This class uses the `scipy.optimize` library to optimize the parameters
specified in the `pySolve_Param` Excel range. `pySolve_Param` must adhere to a specific format
to be used with this class.

Returns
-------
scipy.optimize.OptimizeResult
    The result of the optimization

References
----------
`scipy.optimize`:
    https://docs.scipy.org/doc/scipy/reference/optimize.html
`scipy.optimize.OptimizeResult`:
    https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.OptimizeResult.html

Notes
-----
The `pySolve_Param` range must be a named range scoped to the worksheet with the following row names:
    [0]: active?    -> Y/N denoting if param[i] is active
    [1]: param      -> parameter name
    [2]: val        -> parameter value
    [3]: min        -> min bound on parameter
    [4]: max        -> max bound on parameter
    [5]: obj        -> value of objective function
"""
import datetime
import json
import dill
import xlwings as xw
from pathlib import Path
import numpy as np
from scipy.optimize import minimize, basinhopping, brute, differential_evolution, shgo, dual_annealing, direct
from scipy.optimize import OptimizeResult, show_options
from utils.file_excel import rg_to_dict
from utils.json import cls_to_json, json_to_cls
from utils.file_excel import XW

# CONSTANTS
THIS_DIR = Path(__file__).parent
DATA_DIR = THIS_DIR / 'data'
    
class Excel_Solver:
    """ 
    Solves an optimization problem set up in Excel using the `scipy.optimize` library.
    
    Optimization methods: https://docs.scipy.org/doc/scipy/reference/optimize.html

    - minimize: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.minimize.html
    - basinhopping: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.basinhopping.html
    - brute: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.brute.html
    - differential_evolution: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.differential_evolution.html
    - shgo: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.shgo.html
    - dual_annealing: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.dual_annealing.html
    - direct: https://docs.scipy.org/doc/scipy/reference/generated/scipy.optimize.direct.html
    """
        
    _SOLVER_MAP = dict(active=1, param=2, val=3, min=4, max=5, obj=6)
    # kwargs for different `scipy.optimize` methods
    _KWARGS = dict(bounds=None, args=())
    #_OPTIONS = show_options(solver='minimize', disp=False)
    _OPT_KWARGS = {
        'minimize': dict(x0=None, method=None, jac=None, hess=None, hessp=None, bounds=None, constraints=(), tol=None, callback=None, options=None),
        'basinhopping': dict(x0=None, niter=100, T=1.0, stepsize=0.5, minimizer_kwargs=None, take_step=None, accept_test=None, callback=None, interval=50, 
                             disp=False, niter_success=None, seed=None, target_accept_rate=0.5, stepwise_factor=0.9),
        'brute': dict(ranges=None, Ns=20, full_output=0, finish=None, disp=False, workers=1),
        'differential_evolution': dict(bounds=None, strategy="best1bin", maxiter=1000, popsize=15, tol=0.01, mutation=(0.5, 1), recombination=0.7, 
                                       seed=None, callback=None, disp=False, polish=True, init='latinhypercube', atol=0, updating='immediate',
                                       workers=1, constraints=(), x0=None, integrality=None, vectorized=False),
        'shgo': dict(bounds=None, constraints=None, n=100, iters=1, callback=None, minimizer_kwargs=None, options=None, sampling_method='simplicial', workers=1),
        'dual_annealing': dict(bounds=None, maxiter=1000, minimizer_kwargs=None, initial_temp=5230.0, restart_temp_ratio=2e-05, visit=2.62, accept=-5.0, maxfun=10000000.0, 
                               seed=None, no_local_search=False, callback=None, x0=None),
        'direct': dict(bounds=None, eps=0.0001, maxfun=None, maxiter=1000, locally_biased=True, 
                       f_min=-np.inf, f_min_rtol=0.0001, vol_tol=1e-16, len_tol=1e-06, callback=None),
        'options': dict(maxfev=None, f_min=None, f_tol=None, maxiter=None, maxev=None, maxtime=None, minhgrdint=None, symmetry=None),
    }
    _OPT_METHODS = {
        'global': ['basinhopping', 'brute', 'differential_evolution', 'shgo', 'dual_annealing', 'direct'],
        'local': ['Nelder-Mead', 'Powell', 'CG', 'BFGS', 'Newton-CG', 'L-BFGS-B', 
                  'TNC', 'COBYLA', 'SLSQP', 'trust-constr', 'dogleg', 'trust-ncg', 'trust-exact', 'trust-krylov']
        }
    
    def __init__(self, book:Path, sheet_name:str="project", param_rg_name:str="pySolve_Param", algo_rg_name:str="pySolve_Algo",
                 objective_dict:dict=None)->None:
        """
        Initializes an instance of the class.

        Parameters
        ----------
        book : Path 
            The path to the Excel file.
        sheet_name : str
            The name of the worksheet.
        param_rg_name : str, optional
            The name of the "param" range in the worksheet. Defaults to "pySolve_Param".
        algo_rg_name : str, optional
            The name of the "settings" range in the worksheet. Defaults to "pySolve_Algo".
        objective_dict : dict, optional
            The objective function dictionary, keys('sheet', 'rg'). Defaults to None.
            
        Attributes
        ----------
        x_param : dict {keys: 'val', 'min', 'max'}
            Active `x` parameters.
        x_param_all : dict {keys: 'val', 'min', 'max'}
            All `x` parameters. `x_param` is made from extracting active params from this variable.
        algo_param : dict
            Hyperparameters for the optimization algorithm
        solution : dict
            Contains results from the optimization.
        
        Private Attributes
        ------------------
        _xw : XW class instance
            Contains attributes {book, sheet, ranges}
        _solver_admin : dict {keys: 'DIR', 'check_file', 'terminate_optimization'}
            Contains keys for optimization termination.
            
        Methods
        -------
        optimize() -> Result
            Performs the optimization as specified on the Excel sheet.
        print_solutions() -> None
            Writes `self.solution` info to an Excel sheet.
        write_solution_to_solver_range(idx) -> None:
            Write `x[idx]` to the Excel solver range.
        """
        # xlwings objects
        self.init_xw(book, sheet_name, param_rg_name, algo_rg_name)
        
        # dicts: (solver_admin, x_param, x_param_all, algo_param)
        self.init_param()
        if len(self.x_param['val'])==0:
            raise ValueError("No active parameters found in the `param_rg_name` range.")
        
        # outputs from `self.optimize()` method
        self.solution = dict(result=None, nfev=0, storage_tol=None, nSolutions=None,
                             f=[], x=[], error=[], penalty=[])   # storage for solutions (result, f(x), x)
        self.solution['storage_tol'] = self.algo_param['param'].pop('storage_tol', None)
        
        # f(x), custom objective
        self._xw_f_custom = XW(book, objective_dict['sheet'], [objective_dict['rg']], ['rg']) if objective_dict else None
    
    # INIT auxiliary methods
    def init_param(self)->None:
        """Constructs the attribute(s): `solver_admin, x_param_all, x_param, algo_param`."""
        self.init_solver_admin()
        self.init_x_param()
        self.init_algo_param()
    
    def init_solver_admin(self)->None:
        """Constructs the attribute(s): `solver_admin`."""
        this_file = Path(__file__)
        DIR = this_file.parent
        self._solver_admin = {
            'path': this_file,
            'script_name': f"{DIR.name}/{this_file.name}", 
            'check_file': DIR / 'stop_optimizer.txt', # check-file for terminating optimization
            'terminate_optimization': False #flag to terminate optimization
        }
        
    def init_xw(self, book, sheet_name, param_rg_name, algo_rg_name)->None:
        """Constructs the attribute(s): `xw`."""
        self._xw = XW(book, sheet_name, [param_rg_name, algo_rg_name], ['rg_x', 'rg_algo'])
    
    def init_x_param(self)->None:
        """Constructs the attribute(s): `x_param_all, x_param`."""
        self.x_param_all = self._x_rg_to_dict(self._xw.ranges['rg_x'])
        self.x_param = self._read_x_active(self.x_param_all)        
    
    def init_algo_param(self)->None:
        """Constructs the attribute(s): `algo_param`."""
        self.algo_param = dict(method=None, param=None, objective='default')
        if self._xw.ranges['rg_algo'] is not None:
            algo_params = self._x_rg_to_dict(self._xw.ranges['rg_algo'])
            self.algo_param['method'], self.algo_param['param'] = self._read_algo_params(algo_params)
            self.algo_param['objective'] = self.algo_param['param'].pop('objective', 'default')
            # add `param_x` to `algo['param']`
            self.algo_param['param']['x0'] = self.x_param.get('val', None)          # `x0`, init guess
            self.algo_param['param']['bounds'] = self.x_param.pop('bounds', None)   # bounds on `x`
            self.algo_param['param']['ranges'] = self.x_param.pop('bounds', None)   # bounds on `x` -- for `scipy.optimize.brute()`
        # self.algo_param['param'] = self._filter_kwargs(self.algo_param['method'], self.algo_param['param'])
        
    def _x_rg_to_dict(self, rg:xw.Range) -> dict:
        """Convenience method to convert "Param" range object to dict."""
        return rg_to_dict(rg, isRowHeader=False, isUnit=False, isTrimNone=False, isLowerCaseKey=False)

    def _read_x_active(self, d:dict)->dict:
        """
        Filters the dictionary to include only entries where 'active?' is 'Y'.

        Parameters
        ----------
        d : dict
            A dictionary containing parameter information. Expected keys are
            'active?', 'param', 'val', 'min', and 'max'. Each key has a list of values.

        Returns
        -------
        dict
            A new dictionary with the same structure as `d`, but only containing
            entries where 'active?' is 'Y'. Also adds keys: {'bounds', 'indices'}
        """
        def del_keys(d:dict, keys:list):
            for key in keys:
                d.pop(key, None)

        # Iterate over each entry and add active parameters to `d_active`
        d_active = {key: [] for key in d}
        d_active['indices'] = []    # Add the 'indices' key
        d_active['bounds'] = []     # Add the 'bounds' key
        for i, isActive in enumerate(d['active?']):
            if isActive == 'Y':
                for key in d:
                    d_active[key].append(d[key][i])
                d_active['indices'].append(i)
                d_active['bounds'].append((d['min'][i], d['max'][i]))
        d_active['bounds'] = [(d['min'][i], d['max'][i]) for i, isActive in enumerate(d['active?']) if isActive == 'Y']
        del_keys(d_active, ['active?', 'min', 'max'])
        return d_active
    
    def get_algo_params(self, method:str=None) -> dict:
        """Returns a copy of the `param_algo['param']` attribute."""
        if method is None:
            d = self.algo_param['param']
        else:
            d = self._OPT_KWARGS.get(method)
        return d.copy()

    def set_algo_params(self, method:str=None, param:dict=None)->None:
        """Modifies the `param_algo` dict which is used by the `optimize()` method.
        
        Parameters
        ----------
        method : str, optional
            Optimization method
        param : dict, optional
            keyword arguments to the `scipy.optimize.<method>` method
        """
        d = self.algo_param['param']
        if param is not None:
            param.setdefault('bounds', d.get('bounds', None))
            param.setdefault('x0', d.get('x0', None))
        self.algo_param['method'] = method
        self.algo_param['param'] = param
    
    def _read_algo_params(self, algo_params) -> tuple[str, dict]:
        """
        Reads the algorithm parameters from the Excel sheet and stores them in a dictionary.

        Returns
        -------
        dict
            A dictionary containing algorithm parameters and their values.
        """
        def get_val(d:dict, key:str):
            """Extracts the value from d[key] and converts its type."""
            value = d[key]
            if isinstance(value, str):
                if value.isdigit():
                    return int(value)
                elif self._is_float(value):
                    return float(value)
            return value

        # read `rg` and populate `algo_param`
        algo_param = {}
        row = 1
        for key in algo_params.keys():
            if not key:
                break  # Stop if row is empty
            algo_param[key] = get_val(algo_params, key)
            row += 1

        algo_method = algo_param.pop('algo_method', 'L-BFGS-B')
        try:
            algo_param.update(self._OPT_KWARGS[algo_method])
        except Exception:
            algo_param.update(self._OPT_KWARGS['minimize'])
        return algo_method, algo_param

    def _is_float(self, s:str)->bool:
        """Helper method to check if a string can be converted to a float."""
        try:
            float(s)
            return True
        except ValueError:
            return False
    
    # OPTIMIZATION methods
    def _pass_x_to_solver(self, x)->None:
        """Passes the current values of the optimization parameters to the Excel sheet.

        Parameters
        ----------
        x : list or numpy array
            The current values of the optimization parameters.
        """
        s = self._SOLVER_MAP
        for i, xi in zip(self.x_param['indices'], x):
            self._xw.ranges['rg_x'](s['val'], i+2).value = xi
    
    def _get_objective(self, objective_type:str)->dict or float:
        """Read value from objective cell."""
        s = self._SOLVER_MAP
        if objective_type == 'default':
            f = self._xw.ranges['rg_x'].value[s['obj']-1]
            c = 1 # Assuming the objective value is in col=1
            return dict(f=f[c], error=f[c+1], penalty=f[c+2]) 
        else:
            # read x, y, y*
            data = self.rg_obj
            # gpt: read {x,y,y*} from `data`
            # set 1 is in col{x=1,y=2,y*=3}, 2 is in col{x=4,y=5,y*=6} and so on. 
            
    def _objective_function(self, x, objective_type="default", write_to_storage=True)->float:
        """Reads value of objective from `Solve_Knobs` range.

        Parameters
        ----------
        x : list or numpy array
            The current values of the optimization parameters.

        Returns
        -------
        float
            The value of the objective function read from the Excel sheet.
        """
        if self._is_terminate_optimization(self._solver_admin['check_file']):
            self._solver_admin['terminate_optimization'] = True
            return float(np.inf)
        # Pass active `x` to Excel solver range
        self._xw.app.calculation = 'manual'
        self._pass_x_to_solver(x)
        self._xw.app.calculate()
        self._xw.app.calculation = 'automatic'
        # read or evaluate objective
        obj = self._get_objective(objective_type)
        if isinstance(obj, dict):
            f, error, penalty = obj['f'], obj['error'], obj['penalty']
        else:
            f, error, penalty = obj, None, None
        self.solution['nfev'] += 1  #increment nfev counter
        # store solution
        eps = self.solution['storage_tol']
        if write_to_storage and eps is not None and (f < eps or (error is not None and error < eps)):
            self.solution['f'].append(f)
            self.solution['error'].append(error)
            self.solution['penalty'].append(penalty)
            self.solution['x'].append(x)
        return f

    def _filter_kwargs(self, method:str, kwargs:dict)->dict:
        """
        Filters keyword arguments to include only those that are valid for the given method.

        Parameters
        ----------
        method : str
            The optimization method.
        kwargs : dict
            The full set of keyword arguments.

        Returns
        -------
        dict
            Filtered keyword arguments.
        """
        algo_params = self._OPT_KWARGS.get(method, {})
        valid_params = {k: v for k, v in kwargs.items() if k in algo_params.keys()}
        # Convert specific keys to integers
        int_keys = ['popsize', 'workers']  # Add other keys as needed
        for key in int_keys or 'iter' in key:
            if key in valid_params and not isinstance(valid_params[key], int):
                try:
                    valid_params[key] = int(valid_params[key])
                except ValueError as e:
                    raise ValueError(
                        f"Value for '{key}' must be an integer. Got '{valid_params[key]}'."
                    ) from e

        return valid_params
    
    def _solver_callback(self, xk)->bool:
        """Callback function to stop the solver if a stopping condition is met.

        Parameters
        ----------
        xk : numpy array
            The current solution at iteration k.

        Returns
        -------
        bool
            True if the optimization should terminate, False otherwise.
        """
        # Check for external termination signal
        print("Entered callback", f"xk={xk}")
        if self._is_terminate_optimization():
            print("Termination signal detected. Stopping optimization.")
            return True
        
        # If the current objective is below the threshold, stop the solver
        # obj = self._get_objective(self.algo_method)
        # if obj < self.callback_threshold:
        #     return True
        return False

    def _is_terminate_optimization(self, check_file="stop_optimizer.txt") -> bool:
        """Check for a specific condition or signal to terminate the optimization."""
        return False if check_file is None else Path(check_file).exists()

    def _optimize_args(self, method:str, opt_kwargs:dict)->tuple[str, dict]:
        """Aux method for `optimize()` to handle `opt_args` input."""
        # handle `method` input
        if method is not None:
            self.algo_param['method'] = method
        method = self.algo_param['method'].lower()
        
        # TODO This needs to be handled differently. Should not try to build `kwargs`, rather allow user full control of the definition of `kwargs`. 
        # There are too many details to handle to make this method work with all the different optimization algorithms. Give the user a method for getting the default `kwargs` for a 
        # particular method and then they can modify those `kwargs` how they see fit and then deal with any errors thrown by the optimization algorithms.
        if not opt_kwargs:
            opt_kwargs = self.algo_param['param'] # algorithm parameters
        else:
            opt_kwargs.update(self.algo_param['param'])  # Combine with additional kwargs, if any
        local_minimizer= opt_kwargs.get('local_minimizer', None)
        bounds = opt_kwargs.get('bounds', None)
        opt_kwargs = self._filter_kwargs(method, opt_kwargs)
            
        # handle `minimizer_kwargs` if a key in `kwargs`
        if 'minimizer_kwargs' in opt_kwargs:
            minimizer_kwargs = self._OPT_KWARGS['minimize']
            minimizer_kwargs.pop('x0', None)
            minimizer_kwargs['method'] = local_minimizer
            minimizer_kwargs['bounds'] = bounds
            if opt_kwargs['minimizer_kwargs'] is not None:
                minimizer_kwargs.update(opt_kwargs.pop('minimizer_kwargs', {}))
            opt_kwargs['minimizer_kwargs'] = minimizer_kwargs
        
        if 'x0' in opt_kwargs:
            opt_kwargs['x0'] = self.x_param.get('val', np.ones(len(bounds), dtype=float))

        # add callback to `kwargs` (allows user to stop optimization early)
        if opt_kwargs['callback'] is None:
            opt_kwargs['callback'] = None #self._solver_callback       
            
        return method, opt_kwargs
    
    def _optimize_run(self, method:str, args=(), **opt_kwargs:dict)->OptimizeResult:
        """Aux method for `optimize()` to call optimization routine."""
        try:
            if method == 'differential_evolution':
                result = differential_evolution(self._objective_function, args=args, **opt_kwargs)
            elif method == 'basinhopping':
                result = basinhopping(self._objective_function, **opt_kwargs)
            elif method == 'brute':
                result = brute(self._objective_function, args=args, **opt_kwargs)
            elif method == 'shgo':
                if opt_kwargs['minimizer_kwargs']['method'] is None:
                    opt_kwargs['minimizer_kwargs']['method'] = 'SLSQP'
                result = shgo(self._objective_function, args=args, **opt_kwargs)
            elif method == 'dual_annealing':
                result = dual_annealing(self._objective_function, args=args, **opt_kwargs)
            elif method == 'direct':
                result = direct(self._objective_function, args=args, **opt_kwargs)  
            else: # Assuming the default case is a local minimizer
                result = minimize(self._objective_function, args=args, **opt_kwargs)
        except Exception as e:
            result = None
            if self._solver_admin['terminate_optimization']:
                print("Optimization terminated by stop file!")
            else:
                raise e
        return result
            
    def optimize(self, method:str=None, args=(), **opt_kwargs) -> OptimizeResult:
        """
        Optimizes the parameters specified in the Excel 'pySolve_Param' range using
        the optimization algorithm specified in the Excel 'pySolve_Settings' range.

        The method updates parameter values in the Excel file and evaluates
        the objective function, which should be reflected in the specified Excel range.
        The optimization algorithm is chosen based on the 'method' argument.
        The method supports local and global optimization algorithms.

        Parameters
        ----------
        method : str, optional
            The optimization algorithm to use. It can be any local optimization
            method supported by `scipy.optimize.minimize` (e.g., 'L-BFGS-B', 'SLSQP'),
            'differential_evolution' for global optimization, or 'basinhopping' for
            a stochastic global optimization algorithm.
        args : tuple, optional
            Additional arguments to be passed to the objective function.
        **opt_kwargs
            Additional keyword arguments to be passed to the optimization function.
            For 'basinhopping', you can pass 'minimizer_kwargs' as a dictionary
            to specify arguments for the underlying local optimization method.
            Other arguments are specific to the chosen optimization method.

        Returns
        -------
        result : scipy.optimize.OptimizeResult
            return object from the optimization method.
        This method updates...
            - Excel file with the optimized parameters directly.
            - self.solution['result'] with the `result` object.

        Example
        -------
        >>> solver = Excel_Solver(book=book_path, sheet_name="Sheet1", param_rg_name="pySolve_Param", algo_rg_name="pySolve_Algo")
        >>> solver.optimize(method='basinhopping')
        >>> solver.close_excel()
        """
        # get optimization method and kwargs
        method, opt_kwargs = self._optimize_args(method, opt_kwargs)    
        
        # modify EXCEL app
        # NOTE: I don't know if `screen_updating` is causing the problems with Python crashing. 
        #self.xw.app.screen_updating = False
        
        # run optimizer
        args = ('default', True) #objective_type, write_to_list
        try:
            if method == 'differential_evolution':
                result = differential_evolution(self._objective_function, args=args, **opt_kwargs)
            elif method == 'basinhopping':
                result = basinhopping(self._objective_function, **opt_kwargs)
            elif method == 'brute':
                result = brute(self._objective_function, args=args, **opt_kwargs)
            elif method == 'shgo':
                if opt_kwargs['minimizer_kwargs']['method'] is None:
                    opt_kwargs['minimizer_kwargs']['method'] = 'SLSQP'
                result = shgo(self._objective_function, args=args, **opt_kwargs)
            elif method == 'dual_annealing':
                result = dual_annealing(self._objective_function, args=args, **opt_kwargs)
            elif method == 'direct':
                result = direct(self._objective_function, args=args, **opt_kwargs)  
            else: # Assuming the default case is a local minimizer
                result = minimize(self._objective_function, args=args, **opt_kwargs)
        except Exception as e:
            result = None
            if self._solver_admin['terminate_optimization']:
                print("Optimization terminated by stop file!")
            else:
                raise e
            
        # modify EXCEL app
        #self.xw.app.screen_updating = True
        
        # Update the optimized values in the Excel sheet
        self.solution['nSolutions'] = len(self.solution['f'])
        if result is None:
            if self.solution['nSolutions'] > 0:
                f_min = min(self.solution['f'])
                x = self.solution['x'][self.solution['f'].index(f_min)]
            else:
                x = opt_kwargs['x0']
        else:
            x = result.x
        self._objective_function(x, *args)
        self.solution['result'] = result
        return result
    
    # region - writing to Excel sheet
    def write_solution_to_solver_range(self, idx:int)->None:
        """Write x[idx] to the Excel solver range.
        
        Parameters
        ----------
        idx : int
            idx of the `solution['x']` attribute to print to Excel range.
        """
        # extract solution from `idx`
        f = self.solution['f'][idx]
        x = self.solution['x'][idx]
        
        # write `x` to solver range
        f0 = self._objective_function(x, write_to_storage=False)
        x0 = [f"{v:.5f}" for v in x]
        x1 = [f"{self.x_param['param'][i]}={x0[i]}" for i in range(min(len(self.x_param['param']), len(x0)))]
        print(f"idx={idx}; objective={f0}")
        #print(f"x: {self.x_param['param']}={x0}")
        print(f"x: {x1}")
        

    def print_solutions(self, sheet_name="OptimizeResult", **kwargs) -> None:
        """
        Writes the candidate solutions and their corresponding objective values to a new Excel sheet.

        This method creates a new sheet in the workbook with the specified name and records each candidate
        solution's parameters and its objective function value. The solutions are those that have met
        certain criteria during the optimization process, such as satisfying a tolerance threshold.

        Parameters
        ----------
        sheet_name : str, optional
            The name of the Excel sheet where the solutions will be written. If a sheet with the given name
            already exists, it will be overwritten. The default name is "OptimizeResult".
        **kwargs : dict, optional
            - 'autofit' (bool): To automatically adjust the column widths on the sheet.
            - 'script': The name of the script that created the sheet.

        Notes
        -----
        The method assumes that `self.solution` is a dict containing `f: objective`, `x: param`
        corresponding to the candidate solutions.

        Example
        -------
        >>> solver = Excel_Solver(book=book_path, sheet_name="project")
        >>> result = solver.optimize()  # returns `OptimizeResult` object
        >>> solver.print_solutions(sheet_name="OptimizeResult")
        """
        # Create a new Excel sheet for the solutions (delete the sheet if it already exists)
        if sheet_name in [s.name for s in self._xw.book.sheets]:
            self._xw.book.sheets[sheet_name].delete()
        sheet = self._xw.book.sheets.add(name=sheet_name)

        # Write the (info, initial, final, `result` object) to sheet
        # initial: initial values of the FULL parameter set (including inactive parameters)
        # final: values and properties of the `result` of the optimization
        sol = self.solution
        result = sol['result']
        data = [
            ["info:", "This sheet created using the `Excel_Solver.print_solutions()` method, where the solutions were generated by the `.optimize()` method."],
            ["problem:", "min(f(x)), where `x` is the set of active parameters and `f` is the objective."],
            ["script:", kwargs.get('script', self._solver_admin['script_name'])],
            ["book:", f"{self._xw.book.name}"],
            ["sheet:", f"{self._xw.sheet.name}"],
            ["ranges:", f"{self._xw.ranges['rg_x'].name.name}, {self._xw.ranges['rg_algo'].name.name}"],
            [""],
            ["problem setup:"],
            [f"x[0] and x-bounds (defined in Excel range `{self._xw.ranges['rg_x'].name.name}`)"],
            ["", "objective", "error", "parameters (all)"],
            ["", "", ""] + list(range(len(self.x_param_all['param']))),
            ["", "f(x)", "err(x)"] + self.x_param_all["param"],
            ["initial:", self.x_param_all["obj"][0], self.x_param_all["obj"][1]] + self.x_param_all["val"],
            ["min:", "", ""] + self.x_param_all["min"],
            ["max:", "", ""] + self.x_param_all["max"],
            [""],
            ["", "objective", "error", "parameters (active)"],
            ["", "f(x)", "err(x)"] + self.x_param['param'],
            ["initial:", self.x_param_all["obj"][0], self.x_param_all["obj"][1]] + self.x_param['val'],
        ]
        if result is None:
            if sol['nSolutions'] > 0:
                f_min = min(sol['f'])
                idx = sol['f'].index(f_min)
                e_min = sol['error'][idx]
                x_min = sol['x'][idx]
                data_result = [
                    ["final:", f_min, e_min] + x_min.tolist(),
                ]
            else:
                data_result = [
                    ["optimization failed!"]
                ]
        else:
            data_result = [
                ["final:", result['fun'], ""] + result['x'].tolist(),
                [""],
                ["optimizer:", f"algorithm / hyperparameters (defined in Excel range `{self._xw.ranges['rg_algo'].name.name}`)"],
                ["algo_method:", self.algo_param['method']],
                ["algo_param:"] + [f"{key}={val}" for key, val in self.algo_param['param'].items()],
                [""],
                ["scipy.optimize.OptimizeResult:"],
                ["message:", result['message']],
                ["success:", result['success']],
                ["fun:", result['fun']],
                ["nfev:", result['nfev']],
                ["nit:", result['nit']],
            ]
        data += data_result

        # Writing data to Excel using a loop
        self._xw.app.screen_updating = False
        for i, row_data in enumerate(data, start=1):
            sheet.range(f"A{i}").value = row_data

        # Write the candidate solutions (header, active params, f(x), x)
        if sol['nSolutions'] > 0:
            data = [
                [""],
                ["solutions:", f"all candidate `x` that yield `f(x) < {sol['storage_tol']}`."],
                ["nSolutions:", sol['nSolutions']],
                ["idx", "objective", "error", "parameters (indices / names / values)"],
                ["", "", ""] + self.x_param['indices'],
                ["", "f(x)", "err(x)"] + self.x_param['param']
            ]

            for i, row_data in enumerate(data, start=i+1):
                sheet.range(f"A{i}").value = row_data

            r = i+1
            for i, (f, e, x) in enumerate(zip(sol['f'], sol['error'], sol['x'])):
                sheet.range(f"A{r}").value = [i, f, e] + x.tolist()
                r += 1
        self._xw.app.screen_updating = True

        # Apply `kwargs`
        if kwargs.get('autofit', False):
            sheet.autofit('columns')
    # endregion
    
    # region - file management
    def close_excel(self)->None:
        """Closes the Excel file and releases all associated resources."""
        self._xw.book.save()
        self._xw.book.close()
        self._xw.app.quit()
    
    def _make_file_path(self, file_extension='json')->Path:
        """Aux method for making file path."""
        t = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")    # Get the current date and time to create a unique filename
        path = self._solver_admin['path'].parent
        book = self._xw.book.name.split('.')[0]
        return path / f"{book}_{t}.{file_extension}"
    
    def to_file(self, file_path:Path=None, file_extension:str=None) -> None:
        """Master method for writing instance to file."""
        valid_extensions = ['json', 'pkl']
        def get_path(path)->Path:
            """Returns """
            return self._make_file_path(file_extension) if path is None else Path(path)

        def get_file_extension(path, file_extension)->str:
            if file_extension is None:
                return path.suffix if path.suffix != '' else 'json'
            else:
                return file_extension

        path = get_path(file_path)
        file_extension = get_file_extension(path, file_extension)
        if file_extension in valid_extensions:
            if file_extension == 'json':
                self.to_json(file_path)
            elif file_extension == 'pkl':
                self.to_pickle(file_path)
        else:
            raise ValueError("Invalid file extension! {file_extension}")
    
    def to_pickle(self, file_path:Path=None)->None:
        """Write instance to pickle file.
        Args:
            file_path (Path, optional): Path to pickle file. Defaults to None.
        """
        if file_path is None:
            file_path = self._make_file_path('pkl')
        
        with open(file_path, 'wb') as f:
            dill.dump(self, f)
    
    @staticmethod
    def from_pickle(file_path:Path):
        """Create instance from pickle file.
        Args:
            file_path (Path): Path to pickle file.
        Returns:
            class instance.
        """
        with open(file_path, 'rb') as f:
            return dill.load(f)
    
    def to_json(self, file_path:Path=None, indent=4)->None:
        """Dump public attributes to JSON file."""
        if file_path is None:
            file_path = self._make_file_path('json')
        cls_to_json(self, file_path, indent=indent)

    @classmethod
    def from_json(cls, file_path:Path):
        """Load attributes from JSON file."""
        return json_to_cls(cls, file_path)
        
    # endregion
    
# SCRIPT

def run_solver(book_path:Path, **kwargs):
    """
    Parameters:
        book_path (Path): The path to the book.
        **kwargs: Optional keyword arguments.
        - sheet_name (str): The name of the sheet, default="project".
        - param_rg_name (str): The name of the "parameter" range name, default="pySolve_Param".
        - algo_rg_name (str): The name of the "hyperparameters" range name, default="pySolve_Algo".
    Returns:
        None
    """
    # create instance of solver
    solver = Excel_Solver(book=book_path, sheet_name=kwargs.get("sheet_name", "project"), 
                          param_rg_name=kwargs.get("param_rg_name", "pySolve_Param"), 
                          algo_rg_name=kwargs.get("algo_rg_name", "pySolve_Algo"),
    )
    
    # modify algorithm parameters (methods: None, 'basinhopping', 'differential_evolution', 'shgo', 'dual_annealing', 'direct')
    algo_method = None
    if algo_method:
        opt_params = solver.get_algo_params(method=algo_method)
        opt_params['bounds'] = solver.algo_param['param']['bounds']
        solver.set_algo_params(method=algo_method, param=opt_params)
    
    # use `optimize` method to solve `min(f(x))`
    result = solver.optimize()

    # save instance to file
    file_name = kwargs.get('file_name')
    if file_name is None:
        file_name = f"{solver._xw.book.name.split('.')[0]}"
        file_ext = 'json'    #{json, pkl, None}
        file_path = DATA_DIR / f"{file_name}.{file_ext}"
    else:
        file_path = DATA_DIR / file_name
    if file_ext == 'pkl':
        solver.to_pickle(file_path)
    elif file_ext == 'json':
        solver.to_json(file_path)
    
    # print candidate solutions to sheet
    solver.print_solutions(sheet_name=kwargs.get("solution_sheet"))
    print(f"optimization result: {result}")

def load_solver(file_path:Path, book_path:Path, **kwargs):
    solver = Excel_Solver.from_json(file_path)
    solver.init_xw(book_path, sheet_name=kwargs.get('sheet_name'), param_rg_name=kwargs.get('param_rg_name'), algo_rg_name=kwargs.get('algo_rg_name'))
    solver.init_param()
    solver.print_solutions(sheet_name=kwargs.get('solution_sheet'))
    print(solver)

if __name__ == "__main__":
    
    # set demo path
    demo_path = THIS_DIR / "optimizer_demo.xlsx"

    # set path to Excel book where optimization will occur
    DIR = Path(r"C:\Users\cjsis\Documents\Ennova\Clients\Oxy\Oxy-MP-561-3\fpxl")
    book_path = DIR / "fpxl_Oxy-MP-561-3.xlsm"
    kwargs = dict(
        sheet_name='project', param_rg_name='pySolve_Param', algo_rg_name='pySolve_Algo', 
        solution_sheet='solutions',
        file_name='fpxl_Oxy-MP-561-3_M10-b.json',
    )
    
    run_optimizer = True
    if run_optimizer:
        run_solver(book_path, **kwargs)
    else:
        file_path = DATA_DIR / 'fpxl_Oxy-MP-561-3_M10_2.json'
        load_solver(file_path, book_path, **kwargs)
    print(f"{Path(__file__).name} complete!")
