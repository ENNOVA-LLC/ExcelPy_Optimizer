""" Script to call the `Excel_Solver` class. 

This script performs an optimization using the `Excel_Solver` class in the `optimizer` module.
"""
import numpy as np
from pathlib import Path
from optimizer import Excel_Solver
from utils.config import ROOT_DIR, DATA_DIR
from utils.utils import make_list_range, get_solution_indices

# script path
isRunSolver = False
isPrintSolutions = False
isSaveInstance = False
isCopyFigs = True

# -- set path to Excel book where optimization occurs
DIR = Path(r"C:\Users\cjsis\Documents\Ennova\Clients\Oxy\Oxy-Tiberius")
book_path = DIR / 'fpxl_Oxy-Tiberius.xlsm'
solver_args = dict(
    sheet_name='project_FL1', param_rg_name='pySolve_Param', algo_rg_name='pySolve_Algo',
    solution_sheet='FL1_OptimizeResult',
    file_name='fpxl_Oxy-FL1_2.json',
)
# -- dict for copying figs to `solution` sheet
excel_fig_dict = dict(
    solution_sheet=solver_args['solution_sheet'],
    to_col='H',
    fig_sheet='asim_FL1', fig_name='Group 1'
)

# run or load solver instance
if isRunSolver:
    # create instance of solver
    solver = Excel_Solver(book=book_path, sheet_name=solver_args.get("sheet_name", "project"), 
                        param_rg_name=solver_args.get("param_rg_name", "pySolve_Param"), 
                        algo_rg_name=solver_args.get("algo_rg_name", "pySolve_Algo"), 
                        kwargs=solver_args)
    # use `optimize` method to solve `min(f(x))`
    result = solver.optimize()
else:
    # load instance of solver and print result
    file_name = DATA_DIR / solver_args.get('file_name')
    solver = Excel_Solver.from_json(file_name)
    result = solver.solution['result']

# print optimization results
print(f"optimization result: {solver.solution['result']}")

# print candidate solutions to sheet and write figures
if isPrintSolutions:
    solver.print_solutions(sheet_name=solver_args.get('solution_sheet'))

# save instance to file
if isSaveInstance:
    file_name = solver_args.get('file_name')
    file_path = DATA_DIR / file_name if file_name is not None else None
    solver.to_file(file_path)
    
# copy figures to `solution` sheet
tol = 20.0
if isCopyFigs:
    idx_list = None #make_list_range([(5503, 5507), (5557, 5562), (5585, 5588)])
    solver.copy_figure_to_solution_sheet(solution_tol=tol, idx_list='all', excel_dict=excel_fig_dict)


idx_list = get_solution_indices(solver.solution['f'], solution_tol=tol)
#solver.write_solution_to_solver_range(idx=0)#idx_list[0])
print(f"{Path(__file__).parent.name}/{Path(__file__).name} complete!")
