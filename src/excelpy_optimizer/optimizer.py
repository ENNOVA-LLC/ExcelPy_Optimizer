""" `optimizer`

Optimizes the parameters in the `pySolve_Param` range.

The `optimizer.Excel_Solver` class uses the `scipy.optimize` library to optimize the parameters
specified in the `pySolve_Param` Excel range. `pySolve_Param` must adhere to a specific format
(shown in the `Notes` section) to be used with this class.

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
- The `pySolve_Param` range must be a named range scoped to the worksheet with the following row names:
    [0]: active?    -> Y/N denoting if param[i] is active
    [1]: param      -> parameter name
    [2]: val        -> parameter value
    [3]: min        -> min bound on parameter
    [4]: max        -> max bound on parameter
    [5]: obj        -> value of objective function
- The `pySolve_Algo` range holds the hyperparameters for the chosen optimization algorithm. 
- Refer to the `optimizer_demo.xlsx` file for an example of proper range formatting.

Refer to the `_OPT_KWARGS` attribute to see the default values for the hyperparameters.
"""
import time
import datetime
import json
import tempfile
from PIL import ImageGrab
import dill
import xlwings as xw
from pathlib import Path
import numpy as np
from scipy.optimize import minimize, basinhopping, brute, differential_evolution, shgo, dual_annealing, direct
from scipy.optimize import OptimizeResult, show_options
from .utils.file_excel import rg_to_dict, get_book
from .utils.json import cls_to_json, json_to_cls
from .utils.file_excel import XW

# UTILITY functions
import pywintypes

def robust_excel_command(command_func, max_retries=5, delay=1, last_delay=5):
    """Attempts to execute an Excel command with retries on com_error."""
    delays = [delay] * max_retries
    if last_delay is not None:
        delays += [last_delay]
    for i, delay in enumerate(delays):
        try:
            return command_func()
        except pywintypes.com_error as e:
            print(f"COM error encountered. Attempt {i + 1}/{max_retries}. Retrying in {delay} seconds...")
            time.sleep(delay)
    # Final attempt outside of loop to raise the exception if it fails
    return command_func()

class Excel_Solver:
    """ 
    # Excel-based Optimization Solver
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
    # map for `pySolve_Param` range row numbers
    _PARAM_MAP = dict(active=1, param=2, val=3, min=4, max=5, obj=6)
    # kwargs for different `scipy.optimize` methods
    _KWARGS = dict(bounds=None, args=())
    # `scipy.optimize` methods
    _OPT_METHODS = {
        'global': ['basinhopping', 'brute', 'differential_evolution', 'shgo', 'dual_annealing', 'direct'],
        'local': ['Nelder-Mead', 'Powell', 'CG', 'BFGS', 'Newton-CG', 'L-BFGS-B', 
            'TNC', 'COBYLA', 'SLSQP', 'trust-constr', 'dogleg', 'trust-ncg', 'trust-exact', 'trust-krylov'
        ]
    }
    # default hyperparameters for all `scipy.optimize` methods
    _OPT_KWARGS = {
        'minimize': dict(x0=None, method=None, jac=None, hess=None, hessp=None, bounds=None, constraints=(), tol=None, callback=None, options=None),
        'basinhopping': dict(x0=None, niter=100, T=1.0, stepsize=0.5, minimizer_kwargs=None, take_step=None, accept_test=None, callback=None, interval=50, 
            disp=False, niter_success=None, seed=None, target_accept_rate=0.5, stepwise_factor=0.9
        ),
        'brute': dict(ranges=None, Ns=20, full_output=0, finish=None, disp=False, workers=1),
        'differential_evolution': dict(bounds=None, strategy="best1bin", maxiter=1000, popsize=15, tol=0.01, mutation=(0.5, 1), recombination=0.7, 
            seed=None, callback=None, disp=False, polish=True, init='latinhypercube', atol=0, updating='immediate',
            workers=1, constraints=(), x0=None, integrality=None, vectorized=False
        ),
        'shgo': dict(bounds=None, constraints=None, n=100, iters=1, callback=None, minimizer_kwargs=None, options=None, sampling_method='simplicial', workers=1),
        'dual_annealing': dict(bounds=None, maxiter=1000, minimizer_kwargs=None, initial_temp=5230.0, restart_temp_ratio=2e-05, visit=2.62, accept=-5.0, maxfun=10000000.0, 
            seed=None, no_local_search=False, callback=None, x0=None
        ),
        'direct': dict(bounds=None, eps=0.0001, maxfun=None, maxiter=1000, locally_biased=True, f_min=-np.inf, f_min_rtol=0.0001, vol_tol=1e-16, len_tol=1e-06, callback=None),
        'options': dict(maxfev=None, f_min=None, f_tol=None, maxiter=None, maxev=None, maxtime=None, minhgrdint=None, symmetry=None),
    }
    
    # region - init
    def __init__(self, 
        book:Path, sheet_name:str="project", param_rg_name:str="pySolve_Param", algo_rg_name:str="pySolve_Algo",
        objective_dict:dict=None, external_objective: callable=None, **kwargs
    )->None:
        """
        Initializes an instance of the class.

        Parameters
        ----------
        book : Path 
            The path to the Excel file.
        sheet_name : str or xw.Sheet
            The name of the worksheet.
        param_rg_name : str, optional
            The name of the "param" range in the worksheet. Defaults to "pySolve_Param".
        algo_rg_name : str, optional
            The name of the "settings" range in the worksheet. Defaults to "pySolve_Algo".
        objective_dict : dict, optional
            The objective function dictionary, keys('sheet', 'rg'). Defaults to None.
        external_objective : callable, optional
            The external objective function. Defaults to None.
            
        Attributes
        ----------
        xw : XW class instance
            Contains attributes {book, sheet, ranges}.
        x_param : dict {keys: 'val', 'min', 'max'}
            Active `x` parameters.
        x_param_all : dict {keys: 'val', 'min', 'max'}
            All `x` parameters. `x_param` is made from extracting active params from this variable.
        algo_param : dict
            Hyperparameters for the optimization algorithm
        solution : dict
            Optimization results.
        
        Private Attributes
        ------------------
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
        copy_figure_to_solution_sheet(idx_list) -> None:
            Write figure images to the `solution` sheet corresponding to candidate solutions in `idx_list`.
        """
        # xlwings objects
        self.init_xw(book, sheet_name, param_rg_name, algo_rg_name)
        
        # dicts: (solver_admin, x_param, x_param_all, algo_param)
        self.init_param()
        if len(self.x_param['val'])==0:
            raise ValueError("No active parameters found in the `param_rg_name` range.")
        
        # outputs from `self.optimize()` method
        self.solution = dict(result=None, nfev=0, storage_tol=None, n_solutions=0, idx_min=None,
            f=[], x=[], error=[], penalty=[], sheet=kwargs.get('solution_sheet', 'OptimizeResult')
        )   # storage for solutions (result, f(x), x)
        self.solution['storage_tol'] = kwargs.get('storage_tol', self.algo_param['param'].pop('storage_tol', None))
        
        # f(x), custom objective
        self._xw_f_custom = XW(book, objective_dict['sheet'], [objective_dict['rg']], ['rg']) if objective_dict else None
        self.external_objective = external_objective
    
    # INIT auxiliary methods
    def init_param(self, **kwargs)->None:
        """Constructs the attribute(s): `solver_admin, x_param_all, x_param, algo_param`."""
        self.init_solver_admin(**kwargs)
        if not hasattr(self, 'x_param'):
            self.init_x_param()
            self.init_algo_param()
    
    def init_solver_admin(self, **kwargs) -> None:
        """Constructs the attribute(s): `solver_admin`."""
        this_file = Path(__file__)
        DIR = this_file.parent
        self._solver_admin = {
            'path': this_file,
            'script_name': f"{DIR.name}/{this_file.name}",
            'check_file': DIR / 'stop_optimizer.txt',  # check-file for terminating optimization
            'terminate_optimization': False,  # flag to terminate optimization
        }
        self._solver_admin['storage_path'] = kwargs.get('file_name', self._make_file_path(file_extension='json'))
        
    def init_xw(self, book:xw.Book, sheet_name:str, param_rg_name:str, algo_rg_name:str)->None:
        """Constructs the attribute(s): `xw`."""
        self.xw = XW(book, sheet_name, [param_rg_name, algo_rg_name], ['rg_x', 'rg_algo'])
    
    def init_x_param(self)->None:
        """Constructs the attribute(s): `x_param_all, x_param`."""
        self.x_param_all = self._x_rg_to_dict(self.xw.ranges['rg_x'])
        self.x_param = self._read_x_active(self.x_param_all)        
    
    def init_algo_param(self)->None:
        """Constructs the attribute(s): `algo_param`."""
        self.algo_param = dict(method=None, param=None, objective='default')
        if self.xw.ranges['rg_algo'] is not None:
            algo_params = self._x_rg_to_dict(self.xw.ranges['rg_algo'])
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
    # endregion
    
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
        s = self._PARAM_MAP
        for i, xi in zip(self.x_param['indices'], x):
            self.xw.ranges['rg_x'](s['val'], i+2).value = xi
    
    def _get_objective(self, objective_type:str, return_type=dict)->dict|float:
        """Read value from objective cell."""
        s = self._PARAM_MAP
        if objective_type == 'default':
            f = self.xw.ranges['rg_x'].value[s['obj']-1]
            c = 1 # Assuming the objective value is in col=1
            result = dict(f=f[c], error=f[c+1], penalty=f[c+2]) 
        else:
            # read x, y, y*
            data = self.rg_obj
            # gpt: read {x,y,y*} from `data`
            # set 1 is in col{x=1,y=2,y*=3}, 2 is in col{x=4,y=5,y*=6} and so on. 
        return result if return_type is dict else result['f']
    
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
        try:
            # Pass active `x` to Excel solver range
            robust_excel_command(lambda: setattr(self.xw.app, 'calculation', 'manual'))
            self.xw.app.calculation = 'manual'
            self._pass_x_to_solver(x)
            robust_excel_command(lambda: self.xw.app.calculate())
            #time.sleep(0.1)  # Add a delay (s) to ensure Excel has time to finish calculating
            robust_excel_command(lambda: setattr(self.xw.app, 'calculation', 'automatic'))
        except pywintypes.com_error as e:
            print()
            raise e
        
        # read or evaluate objective (f=objective, error=err(y, ys), penalty=constraint violations)
        if self.external_objective is not None:
                # Call the user-supplied function directly
                f = self.external_objective(x)
                error = None
                penalty = None
        else:
            obj = self._get_objective(objective_type)
            if isinstance(obj, dict):
                f, error, penalty = obj['f'], obj['error'], obj['penalty']
            else:
                f, error, penalty = obj, None, None
        
        # store solution
        self.solution['nfev'] += 1  #increment nfev counter
        eps = self.solution['storage_tol']
        if write_to_storage and eps is not None and (f < eps or (error is not None and error < eps)):
            self.solution['n_solutions'] += 1
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

    def _optimize_args(self, method:str, opt_kwargs:dict) -> tuple[str, dict]:
        """Aux method for `optimize()` to handle `opt_args` input."""
        # handle `method` input
        if method is not None:
            self.algo_param['method'] = method
        method = self.algo_param['method'].lower()

        # TODO This needs to be handled differently. Should not try to build `kwargs`, rather allow user full control of the definition of `kwargs`. 
        # There are too many details to handle to make this method work with all the different optimization algorithms. Give the user a method for getting the default `kwargs` for a 
        # particular method and then they can modify those `kwargs` how they see fit and then deal with any errors thrown by the optimization algorithms.
        if not opt_kwargs:
            opt_kwargs = self.algo_param['param']       # Algorithm parameters
        else:
            opt_kwargs.update(self.algo_param['param']) # Combine with additional kwargs, if any
        bounds = opt_kwargs.get('bounds', None)
        opt_kwargs = self._filter_kwargs(method, opt_kwargs)

        # handle `minimizer_kwargs` if a key in `kwargs`
        if 'minimizer_kwargs' in opt_kwargs:
            self._optimize_args_minimizer(bounds, opt_kwargs)
        if 'x0' in opt_kwargs:
            opt_kwargs['x0'] = self.x_param.get('val', np.ones(len(bounds), dtype=float))

        # add callback to `kwargs` (allows user to stop optimization early)
        if opt_kwargs['callback'] is None:
            opt_kwargs['callback'] = None #self._solver_callback       

        return method, opt_kwargs

    def _optimize_args_minimizer(self, bounds, opt_kwargs:dict)->None:
        """Aux method for `optimize_args()` to handle `minimizer_kwargs`."""
        minimizer_kwargs = self._OPT_KWARGS['minimize']
        minimizer_kwargs.pop('x0', None)
        minimizer_kwargs['method'] = opt_kwargs.get('local_minimizer')
        minimizer_kwargs['bounds'] = bounds
        if opt_kwargs['minimizer_kwargs'] is not None:
            minimizer_kwargs.update(opt_kwargs.pop('minimizer_kwargs', {}))
        opt_kwargs['minimizer_kwargs'] = minimizer_kwargs
    
    def _optimize_run(self, method:str, args=(), **opt_kwargs:dict) -> OptimizeResult:
        """Aux method for `optimize()` to call optimization routine."""
        try:
            if method == 'basinhopping':
                result = basinhopping(self._objective_function, **opt_kwargs)
            elif method == 'brute':
                result = brute(self._objective_function, args=args, **opt_kwargs)
            elif method == 'differential_evolution':
                result = differential_evolution(self._objective_function, args=args, **opt_kwargs)
            elif method == 'direct':
                result = direct(self._objective_function, args=args, **opt_kwargs)
            elif method == 'dual_annealing':
                result = dual_annealing(self._objective_function, args=args, **opt_kwargs)
            elif method == 'shgo':
                if opt_kwargs['minimizer_kwargs']['method'] is None:
                    opt_kwargs['minimizer_kwargs']['method'] = 'SLSQP'
                result = shgo(self._objective_function, args=args, **opt_kwargs)
            else:
                result = minimize(self._objective_function, args=args, **opt_kwargs)
        except Exception as e:
            result = None
            if self._solver_admin['terminate_optimization']:
                print("Optimization terminated by stop file!")
            else:
                raise e
        return result
            
    def optimize(self, method:str=None, args:tuple=None, **opt_kwargs) -> OptimizeResult:
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
        >>> solver = Excel_Solver(book=book_path, sheet_name="OptimizeResult", param_rg_name="pySolve_Param", algo_rg_name="pySolve_Algo")
        >>> solver.optimize(method='basinhopping')
        >>> solver.close_excel()
        """
        def check_x0(x0, bounds):
            x0_new = []
            for i, value in enumerate(x0):
                lbound, ubound = bounds[i]
                if value < lbound or value > ubound:
                    # Value is outside bounds, set it to the midpoint
                    midpoint = (lbound + ubound) / 2
                    x0_new.append(midpoint)
                else:
                    x0_new.append(value)
            return x0_new
        
        # get optimization method and kwargs
        method, opt_kwargs = self._optimize_args(method, opt_kwargs)    
        opt_kwargs['x0'] = check_x0(opt_kwargs['x0'], opt_kwargs['bounds'])
        
        # modify EXCEL app
        # NOTE: I don't know if `screen_updating` is causing the problems with Python crashing. 
        #self.xw.app.screen_updating = False

        # run optimizer
        if args is None:
            args = ('default', True) #objective_type, write_to_list
        result = self._optimize_run(method, args, **opt_kwargs)

        # modify EXCEL app
        #self.xw.app.screen_updating = True

        # Store results in `solution` dict
        f = self.solution['f']
        self.solution['result'] = result
        self.solution['n_solutions'] = len(f)
        if self.solution['n_solutions'] > 0:
            idx_min = f.index(min(f))
            self.solution['idx_min'] = idx_min   
        
        # Update the optimized values in the Excel sheet
        if result is not None:
            x = result.x
        elif idx_min > 0:
            x = self.solution['x'][idx_min]
        else:
            x = opt_kwargs['x0']
        self._objective_function(x, *args)
        return result
    
    # region - writing to Excel sheet
    def _get_index_of_min_objective(self)->int:
        """Returns the index of the minimum of the objective function."""
        f = self.solution['f']
        return f.index(min(f))

    def write_solution_to_solver_range(self, idx:int=None, isPrint=True)->None:
        """Write x[idx] to the Excel solver range.
        
        Parameters
        ----------
        idx : int or array-like[int]
            idx of the `solution['x']` attribute to print to Excel range.
            if None, then `idx` will correspond to `min(solution['f'])`.
        """
        # Ensure idx is a list for uniform processing
        if idx is None:
            idx = self._get_index_of_min_objective()
        if not isinstance(idx, (list, tuple, np.ndarray)):
            idx = [idx]
            
        # extract solution from `idx`
        for j in idx:
            f = self.solution['f'][j]
            x = self.solution['x'][j]
            
            # write `x` to solver range
            f0 = self._objective_function(x, write_to_storage=False)
            if isPrint:
                x0 = [f"{xi:.5f}" for xi in x]
                x1 = [f"{self.x_param['param'][i]}={x0[i]}" for i in range(min(len(self.x_param['param']), len(x0)))]
                print(f"idx={j}; objective (from storage)={f}; objective (current)={f0}")
                print(f"x: {x1}")
                
    def copy_figure_to_solution_sheet(self, solution_tol:float=None, idx_list:list[int]=None, excel_dict:dict=None, isPrint=False) -> None:
        """Writes all figures to the solution sheet.
        
        Parameters
        ----------
        solution_tol : float
            The solution tolerance. All figures corresponding to `solutions['f'] < tol` are printed to `solution_sheet`
        idx_list : int or list[int] or str
            If provided, takes precedence over `solution_tol` to build sub-list of `x` solutions.
            If `all` then prints all candidate solutions.
        excel_dict : dict
            A dictionary of args that define the settings for passing the figure from source (`fig_sheet`) to destination (`solution_sheet`).
            - 'book' (str or Path): workbook containing `solution_sheet` (default=self._xw.book)
            - 'solution_sheet' (str): destination sheet for figures (default='OptimizeResult')
            - 'fig_sheet' (str or xw.Sheet): source sheet for figure (default=ActiveSheet)
            - 'fig_name' (str): name of source figure (default='Group 1')
            - 'to_col' (str): destination col for figure (default='H')
            - 'to_row' (int): destination row for figure (default=41)
        isPrint : bool
            If True, prints the solution info to the console
        """
        def copy_fig(fig_dict:dict, max_retries:int=3, sleep_interval:float=5.0)->None:
            """Create a copy of the figure referenced by `fig_dict`."""
            for i in range(max_retries):
                try:
                    fig = fig_dict['fig']
                    fig.api.Copy()
                    return  # Successful copy
                except Exception as e:
                    if i < max_retries - 1:
                        time.sleep(sleep_interval)
                        fig_dict['fig'] = fig_dict['sheet'].shapes[fig_dict['name']]
                    else:
                        raise e # Re-raise the last exception after all retries have failed

        # Define xs, subset of `x` that meets desired `solution_tol`
        if idx_list is not None:
            if isinstance(idx_list, int):
                idx_list = [idx_list]
            elif isinstance(idx_list, str):
                idx_list = list(range(len(self.solution['f'])))
            xs = [self.solution['x'][i] for i in idx_list]
        elif solution_tol is not None:
            idx_list, xs = [], []
            for i, (x, f) in enumerate(zip(self.solution['x'], self.solution['f'])):
                if f < solution_tol:
                    idx_list.append(i)
                    xs.append(x)
            #idx_list = [i for i, f in enumerate(self.solution['f']) if f < solution_tol]
            #xs = [self.solution['x'][i] for i in idx_list]
        else:
            raise ValueError("Either `solution_tol` or `idx_list` must be provided.")

        # `sol_dict` holds reference to `sol['sheet']`, `to_cell` info
        # `sol['sheet']` can be in a separate workbook, must provide excel_dict['book'] if so
        book = get_book(excel_dict.get('book', self.xw.book))
        solution_sheet = excel_dict.get('solution_sheet', self.solution.get('sheet', 'OptimizeResult'))
        if isinstance(solution_sheet, str):
            solution_sheet = book.sheets[solution_sheet]
        elif not isinstance(solution_sheet, xw.Sheet):
            raise TypeError("`solution_sheet` must be of type `str` or `xw.Sheet`.")
        sol_dict = dict(sheet=solution_sheet, to_col=excel_dict.get('to_col', 'H'), to_row=excel_dict.get('to_row', 41))
        
        # `fig_dict` holds reference to fig sheet, name of fig, and figure
        fig_sheet = excel_dict.get('fig_sheet', self.xw.book.api.ActiveSheet)
        fig_name = excel_dict.get('fig_name', 'Group 1')
        if isinstance(fig_sheet, str):
            fig_sheet = self.xw.book.sheets[fig_sheet]
        elif not isinstance(fig_sheet, xw.Sheet):
            raise TypeError("`fig_sheet` must be of type `str` or `xw.Sheet`.")
        fig_dict = dict(sheet=fig_sheet, name=fig_name, fig=fig_sheet.shapes[fig_name])
        
        valid_paste_types = ['image', 'normal']
        paste_type = excel_dict.get('paste_type', valid_paste_types[0])
        if paste_type not in valid_paste_types:
            raise ValueError(f"Invalid `paste_type` value. Must be in set {valid_paste_types}. Got '{paste_type}' instead.")
        
        # Loop over all `idx_list`, pass xs[i] to sheet, copy resultant figure, and paste using info from `sol_dict`.
        screen_updating = self.xw.app.screen_updating
        self.xw.app.screen_updating = True
        for idx, x in zip(idx_list, xs):            
            # Update Excel with solution
            self.write_solution_to_solver_range(idx, isPrint=isPrint)
            
            # Copy source figure
            copy_fig(fig_dict, max_retries=3, sleep_interval=5.0)
            
            # Paste figure to destination cell
            to_cell = sol_dict['sheet'].range(f"{sol_dict['to_col']}{sol_dict['to_row']+idx}")
            if paste_type == 'image':
                # creates a temp file, saves fig to file, writes fig to sheet, then deletes the temp file
                temp_path = Path(tempfile.gettempdir()) / f'fig_{idx}.png'
                img = ImageGrab.grabclipboard()
                img.save(temp_path, 'PNG')
                sol_dict['sheet'].pictures.add(str(temp_path), name=f"fig_{idx}", top=to_cell.top, left=to_cell.left)
                temp_path.unlink()
            elif paste_type == 'normal':
                # this method doesn't work properly because the plot continues referencing the source data, which is updating
                sol_dict['sheet'].api.Paste(to_cell.api)
        # restore
        self.xw.app.screen_updating = screen_updating
        
    def print_solutions(self, sheet_name="OptimizeResult", **kwargs) -> None:
        """
        Writes the candidate solutions and their corresponding objective values to a new Excel sheet.

        This method creates a new sheet in the workbook with the specified name and records each candidate
        solution's parameters and its objective function value. The solutions are those that have met
        certain criteria during the optimization process, such as satisfying a tolerance threshold.

        Parameters
        ----------
        sheet_name : str or xw.Sheet, optional
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
        def make_new_sheet(book:xw.Book, base_sheet_name:str)->xw.Sheet:
            """Create a new `xw.Sheet` object with `base_sheet_name`."""
            if base_sheet_name in [s.name for s in book.sheets]:
                i = 1
                sheet_name = f"{base_sheet_name}_{i}"
                while sheet_name in [s.name for s in book.sheets]:
                    i += 1
                    sheet_name = f"{base_sheet_name}_{i}"
            else:
                sheet_name = base_sheet_name

            # Add the new sheet with the determined name
            sheet = book.sheets.add(name=sheet_name)
            return sheet

        # Create a new Excel sheet for the solutions
        if sheet_name is None:
            sheet_name = self.solution.get('sheet', 'OptimizeResult')
        if isinstance(sheet_name, xw.Sheet):
            book = sheet_name.book
            sheet_name = sheet_name.name
        else:
            book = self.xw.book        
        sheet = make_new_sheet(book, base_sheet_name=sheet_name)
        self.solution['sheet'] = sheet.name
        
        # Write to `sheet`: (info, initial, final, `result` object)
        # initial: 
        # - algorithm hyperparameters (`algo_param`)
        # - initial values of the FULL parameter set (including inactive parameters, `x_param_all`)
        # final: 
        # - values and properties of the `result` of the optimization
        # - all candidate solutions that meet `solution_tol`
        sol = self.solution
        result = sol['result']
        data = [
            ["info:", "This sheet created using the `Excel_Solver.print_solutions()` method, where the solutions were generated by the `.optimize()` method."],
            ["problem:", "min(f(x)), where `x` is the set of active parameters and `f` is the objective."],
            ["script:", kwargs.get('script', self._solver_admin['script_name'])],
            ["book:", f"{self.xw.book.name}"],
            ["sheet:", f"{self.xw.sheet.name}"],
            ["ranges:", f"{self.xw.ranges['rg_x'].name.name}, {self.xw.ranges['rg_algo'].name.name}"],
            ["storage:", f"{self._solver_admin['storage_path'].resolve()}"],
            [""],
            ["problem setup:"],
            ["optimizer:", f"algorithm / hyperparameters (defined in Excel range `{self.xw.ranges['rg_algo'].name.name}`)"],
            ["algo_method:", self.algo_param['method']],
            ["algo_param:"] + [f"{key}={val}" for key, val in self.algo_param['param'].items()],
            [""],
            [f"x[0] and x-bounds (defined in Excel range `{self.xw.ranges['rg_x'].name.name}`)"],
            ["", "objective", "error", "parameters (all)"],
            ["indices:", "", ""] + list(range(len(self.x_param_all['param']))),
            ["", "f(x)", "err(x)"] + self.x_param_all["param"],
            ["initial:", self.x_param_all["obj"][0], self.x_param_all["obj"][1]] + self.x_param_all["val"],
            ["min:", "", ""] + self.x_param_all["min"],
            ["max:", "", ""] + self.x_param_all["max"],
            [""],
            ["results:"],
            ["", "objective", "error", "parameters (active)"],
            ["indices:", "", ""] + self.x_param['indices'],
            ["", "f(x)", "err(x)"] + self.x_param['param'],
            ["initial:", self.x_param_all["obj"][0], self.x_param_all["obj"][1]] + self.x_param['val'],
        ]
        # extract output from the `result` object
        f = sol['f']
        sol['n_solutions'] = len(f)
        
        # if `result` is None, then build `result` from `sol` object
        if sol['n_solutions'] > 0:
            sol['idx_min'] = f.index(min(f))
            i = sol['idx_min']
            f_min, e_min, x_min = sol['f'][i], sol['error'][i], sol['x'][i]
        else:
            f_min = e_min = x_min = None
            
        if result is None and f_min is not None:
            result = dict(fun=f_min, x=x_min, message='`result` object constructed in post-processing', 
                          success=False, nfev=None, nit=None)
        
        try:
            x_list = result['x'].tolist() if 'x' in result and result['x'] is not None else ['N/A']
            data_result = [
                ["final:", result.get('fun', 'N/A'), ''] + x_list,
                [""],
                ["scipy.optimize.OptimizeResult:"],
                ["message:", result.get('message', 'N/A')],
                ["success:", result.get('success', 'N/A')],
                ["fun:", result.get('fun', 'N/A')],
                ["nfev:", result.get('nfev', 'N/A')],
                ["nit:", result.get('nit', 'N/A')],
            ]
        except Exception:
            data_result = [
                ["Error in `result` object!"],
                [''], [''], [''], [''], [''], [''], [''],
            ]
        data += data_result

        # Write `data` to Excel sheet
        self.xw.app.screen_updating = False
        for i, row_data in enumerate(data, start=1):
            sheet.range(f"A{i}").value = row_data

        # Write the candidate solutions (header, active params, f(x), x)
        if sol['n_solutions'] > 0:
            data = [
                ["solutions:", f"all candidate `x` that yield `f(x) < {sol.get('storage_tol', 'storage_tol')}`."],
                ["n_solutions:", sol['n_solutions']],
                ["idx", "objective", "error", "parameters (indices / names / values)"],
                ["", "", ""] + self.x_param['indices'],
                ["", "f(x)", "err(x)"] + self.x_param['param']
            ]
            data += [
                [i, f, e] + x.tolist() for i, (f, e, x) in enumerate(zip(sol['f'], sol['error'], sol['x']))
            ]
            # Write `data` to Excel sheet
            for i, row_data in enumerate(data, start=i+2):
                sheet.range(f"A{i}").value = row_data
        self.xw.app.screen_updating = True

        # Apply `kwargs` to format `sheet`
        if kwargs.get('autofit', False):
            sheet.autofit('columns')
    # endregion
    

    # region - file management
    def close_excel(self)->None:

        """Closes the Excel file and releases all associated resources."""
        self.xw.book.save()
        self.xw.book.close()
        self.xw.app.quit()
    
    def _make_file_path(self, file_extension='json')->Path:
        """Aux method for making file path."""
        t = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M")    # Get the current date and time to create a unique filename
        path = self._solver_admin['path'].parent
        book = self.xw.book.name.split('.')[0]
        return path / f"{book}_{t}.{file_extension}"
    
    def to_file(self, file_path:Path=None, file_extension:str=None) -> None:
        """Master method for writing instance to file.
        
        Parameters
        ----------
        file_path : str or Path, optional
            Path to write instance. If not provided, defaults to book name.
        file_extension : str, optional
            Writes file as `file_extension` type. Overwrites type specified by `file_path`. Valid args: {'pkl', 'json'}
        """
        def get_path(path)->Path:
            """Returns `Path` object where file is to be written."""
            return self._solver_admin.get('storage_path') if path is None else Path(path)

        def get_file_extension(path, file_extension)->str:
            """Returns file extension type."""
            if file_extension is not None:
                return file_extension
            return path.suffix[1:] if path.suffix else 'json'                

        valid_extensions = ['json', 'pkl']
        path = get_path(file_path)
        file_extension = get_file_extension(path, file_extension)
        if file_extension not in valid_extensions:
            raise ValueError(f"Invalid file extension: `{file_extension}`! Valid extensions include: {valid_extensions}")

        full_path = path.with_suffix(f'.{file_extension}')
        if 'json' in file_extension:
            self._to_json(full_path)
        elif 'pkl' in file_extension:
            self._to_pickle(full_path)
    
    def _to_pickle(self, file_path:Path)->None:
        """Write instance to pickle file.
        Args:
            file_path (Path, optional): Path to pickle file. Defaults to None.
        """
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
    
    def _to_json(self, file_path:Path, indent=4)->None:
        """Dump public attributes to JSON file."""
        cls_to_json(self, file_path, indent=indent)

    @classmethod
    def from_json(cls, file_path:Path):
        """Load attributes from JSON file."""
        instance = json_to_cls(cls, file_path)
        instance.init_param()
        return instance
        
    # endregion
    
# SCRIPT
if __name__ == "__main__":
    print("This module holds the class `Excel_Solver`. To see this class in use, refer to the file `main.py`.")
