""" Script to call the `Excel_Solver` class. 

This script performs an optimization using the `Excel_Solver` class in the `optimizer.py` module.

Instructions
------------
1. Import simulation settings.
    - Create a `specs` file in the `specs` folder.
    - Use `import` statement to import the `settings` dict from this file.
2. Define script path.
    - Set `run_method` which specifies the the `run_settings` dict. 
    - This defines which operations in the script are executed.
3. Run this script.
    - The following operations are executed given the `run_settings` dict.
        - a. Run or Load `solver` instance.
        - b. Print candidate solutions to `solution` sheet.
        - c. Save `solver` instance to file.
        - d. Copy figures to `solution` sheet.
        - e. Write a specific solution to the solver range
"""
import numpy as np
from pathlib import Path
from optimizer import Excel_Solver
from utils.config import ROOT_DIR, DATA_DIR
from utils.utils import make_list_range, get_solution_indices

# Step 1: Import simulation settings
from specs.MP561_u import specs_M10 as specs
#from specs.Tiberius_u import specs_FL10 as specs

# Step 2: Define script path
# the `run_settings` dict specifies which operations in the script are executed
run_method = 'solver'
run_settings = {
    "solver": dict(isRunSolver=True, isPrintSolutions=True, isSaveInstance=True, isCopyFigs=False),
    "figure": dict(isRunSolver=False, isPrintSolutions=False, isSaveInstance=False, isCopyFigs=True),
    "write_solution": dict(isRunSolver=False, isPrintSolutions=False, isSaveInstance=False, isCopyFigs=False, isWriteSolution=True)
}
run_settings = run_settings[run_method]

# extract from `settings`
excel_fig_dict = specs.pop('excel_fig_dict')

# Step 3: Perform operations
# 3a) run or load SOLVER instance
if run_settings['isRunSolver']:
    # create instance of solver then use `optimize()` method to solve `min(f(x))`
    solver = Excel_Solver(book=specs.pop('book_path'), 
                          sheet_name=specs.get("sheet_name", "project"), 
                          param_rg_name=specs.get("param_rg_name", "pySolve_Param"), 
                          algo_rg_name=specs.get("algo_rg_name", "pySolve_Algo"), 
                          kwargs=specs)
    result = solver.optimize()
else:
    # load instance of solver and print result
    file_name = DATA_DIR / specs.get('file_name')
    solver = Excel_Solver.from_json(file_name)
    result = solver.solution['result']

# print optimization results
print(f"optimization result: {solver.solution['result']}")

# 3b) Print candidate solutions to sheet
if run_settings['isPrintSolutions']:
    solver.print_solutions(sheet_name=specs.get('solution_sheet'))

# 3c) Save instance to file
if run_settings['isSaveInstance']:
    file_name = specs.get('file_name')
    file_path = DATA_DIR / file_name if file_name is not None else None
    solver.to_file(file_path)
    
# 3d) Copy figures to `solution` sheet
if run_settings['isCopyFigs']:
    tol = excel_fig_dict.pop('tol', None)
    if tol is not None:
        idx_list = get_solution_indices(solver.solution['f'], solution_tol=tol)
    else:
        idx_list = excel_fig_dict.pop('idx_list', None) #make_list_range([(5503, 5507), (5557, 5562), (5585, 5588)])
    solver.copy_figure_to_solution_sheet(solution_tol=tol, idx_list=idx_list, excel_dict=excel_fig_dict)

# 3e) Write a specific solution to the solver range
if run_settings.get('isWriteSolution', False):
    idx = excel_fig_dict.get('solution_idx', None)
    solver.write_solution_to_solver_range(idx=idx)
    
print(f"{Path(__file__).parent.name}/{Path(__file__).name} complete!")
