from pathlib import Path
from optimizer import Excel_Solver

# CONSTANTS
THIS_DIR = Path(__file__).parent
DATA_DIR = THIS_DIR / 'data'

# FUNCTIONS
def run_solver(book_path:Path, **kwargs)->None:
    """Builds instance of `Excel_Solver` class and runs the `.optimize()` method.
    
    Parameters:
        book_path (Path): The path to the book.
        **kwargs: Optional keyword arguments.
        - sheet_name (str): The name of the sheet, default="project".
        - param_rg_name (str): The name of the "parameter" range name, default="pySolve_Param".
        - algo_rg_name (str): The name of the "hyperparameters" range name, default="pySolve_Algo".
        - file_name (Path): The name of the file to save the results to.
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
    file_ext = 'json'
    if file_name is None or isinstance(file_name, Path):
        file_path = file_name
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
    """Loads instance of `Excel_Solver` class from file and prints solutions to Excel sheet."""
    solver = Excel_Solver.from_json(file_path)
    solver.init_xw(book_path, sheet_name=kwargs.get('sheet_name'), param_rg_name=kwargs.get('param_rg_name'), algo_rg_name=kwargs.get('algo_rg_name'))
    solver.init_param()
    solver.print_solutions(sheet_name=kwargs.get('solution_sheet'), script=kwargs.get('script'))
    print(solver)

# SCRIPT
if __name__ == "__main__":
    
    # set demo path
    book_path = THIS_DIR / "optimizer_demo.xlsx"

    # set path to Excel book where optimization will occur
    kwargs = dict(
        sheet_name='project', param_rg_name='pySolve_Param', algo_rg_name='pySolve_Algo', 
        solution_sheet='solutions', script=Path(__file__),
        file_name='demo.json',
    )
    
    # run or load solver
    run_optimizer = True
    run_solver(book_path, **kwargs)
    load_solver(DATA_DIR / 'demo.json', book_path, **kwargs)
    
    # designate script completion
    print(f"{Path(__file__).parent.name}/{Path(__file__).name} complete!")