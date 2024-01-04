""" _summary_

_extended_summary_
"""
from pathlib import Path
from optimizer import Excel_Solver

def main(book_path:Path, **kwargs):
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
    
    # modify algorithm parameters
    algo_method = None #[None, 'basinhopping', 'differential_evolution', 'shgo', 'dual_annealing', 'direct']
    if algo_method:
        opt_params = solver.get_algo_params(method=algo_method)
        opt_params['bounds'] = solver.algo_param['param']['bounds']
        solver.set_algo_params(method=algo_method, param=opt_params)
    
    # use `optimize` method to solve
    result = solver.optimize()
    print(result)

    # print candidate solutions to sheet
    solver.print_solutions(sheet_name=kwargs.get("solution_sheet"))

if __name__ == "__main__":
    
    # set demo path
    THIS_DIR = Path(__file__).parent
    demo_path = THIS_DIR / "optimizer_demo.xlsx"

    # set path to Excel book where optimization will occur
    DIR = Path(r"C:\Users\cjsis\Documents\Ennova\Clients\Oxy\Oxy-MP-561-3\fpxl")
    book_path = DIR / "fpxl_Oxy-MP-561-3.xlsm"

    main(book_path, solution_sheet="solutions")