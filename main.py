""" Script to call the `Excel_Solver` class. 

This script performs an optimization using the `Excel_Solver` class in the `optimizer` module.
"""
import numpy as np
from pathlib import Path
from optimizer import Excel_Solver
from utils.config import ROOT_DIR, DATA_DIR

# Define functions to `solve` optimization and `load` optimization results
def run_solver(book_path:Path, **kwargs)->Excel_Solver:
    """
    Args:
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
    return solver

def load_solver(file_path:Path, book_path:Path=None, **kwargs)->Excel_Solver:
    """Loads an instance of the Excel_Solver class from a JSON file and initializes it with the specified book path.

    Args:
        file_path (Path): The path to the JSON file.
        book_path (Path, optional): The path to the Excel book. Defaults to None.
        **kwargs: Optional keyword arguments.
            - sheet_name (str): The name of the sheet. Defaults to None.
            - param_rg_name (str): The name of the "parameter" range name. Defaults to None.
            - algo_rg_name (str): The name of the "hyperparameters" range name. Defaults to None.

    Returns:
        Excel_Solver: An instance of the Excel_Solver class.
    """
    solver = Excel_Solver.from_json(file_path)
    if not hasattr(solver, 'xw'):
        solver.init_xw(book_path, sheet_name=kwargs.get('sheet_name'), param_rg_name=kwargs.get('param_rg_name'), algo_rg_name=kwargs.get('algo_rg_name'))
    return solver

def make_list(start, end, dx=1)->list:
    """Returns list of numbers from `start` to `end`."""
    return np.arange(start, end+dx, dx).tolist()

def make_list_range(ranges:list[list[float]])->list:
    """Returns list from list of range lists."""
    result = []
    for start, end in ranges:
        result += list(range(start, end + 1))
    return result
    
if __name__ == "__main__":
    
    # Define optimizer settings
    
    # set path to Excel book where optimization occurs
    book_path = ROOT_DIR / "optimizer_demo.xlsx"
    book_path = Path(r"C:\Users\cjsis\Documents\Ennova\Clients\Oxy\Oxy-MP-561-3\fpxl\fpxl_Oxy-MP-561-3.xlsm")
    kwargs = dict(
        sheet_name='project', param_rg_name='pySolve_Param', algo_rg_name='pySolve_Algo',
        solution_sheet='OptimizeResult',
        file_name='fpxl_Oxy-MP-561-3_M10_6.json',
    )
    excel_fig_dict = dict(
        book=Path(r'C:\Users\cjsis\Documents\Ennova\Clients\Oxy\Oxy-MP-561-3\fpxl\Oxy-MP-561-3_M10.xlsx'),
        solution_sheet='OptimizeResult', #kwargs['solution_sheet']
        to_col='G',
        fig_sheet='asim', fig_name='Group 1'
    )
    
    # script path
    isRunSolver = True
    isPrintSolutions = True
    isCopyFigs = False
    isSaveInstance = True

    # run or load solver instance
    if isRunSolver:
        solver = run_solver(book_path, **kwargs)
    else:
        solver = load_solver(DATA_DIR / kwargs.get('file_name'), book_path, **kwargs)
    
    # print optimization results
    print(f"optimization result: {solver.solution['result']}")
    
    # print candidate solutions to sheet and write figures
    if isPrintSolutions:
        solver.print_solutions(sheet_name=kwargs.get('solution_sheet'))
    
    # copy figures to `solution` sheet
    if isCopyFigs:
        idx_list = make_list_range([(5503, 5507), (5557, 5562), (5585, 5588)])
        solver.copy_figure_to_solution_sheet(solution_tol=50.0, idx_list=idx_list, excel_dict=excel_fig_dict)
    
    # save instance to file
    if isSaveInstance:
        file_name = kwargs.get('file_name')
        file_path = DATA_DIR / file_name if file_name is not None else None
        solver.to_file(file_path)
        
    solver.write_solution_to_solver_range(idx=44)
    print(f"{Path(__file__).parent.name}/{Path(__file__).name} complete!")
