# %% [markdown]
# # `Excel_Solver` demo
# 
# Jupyter notebook to use `Excel_Solver` class to link Excel to the `scipy.optimize` package. 
# 
# The `Excel_Solver` class implements the minimization algorithms from the `scipy.optimize` package and provides an Excel interface for building the `solver` instance of the `Excel_Solver` class.
# 
# The main methods of the `Excel_Solver` class are:
# + `optimize()`: runs the optimization algorithm.
# + `print_solutions()`: prints the solutions to an Excel sheet.

# %%
# import packages
from pathlib import Path
from excelpy_optimizer import Excel_Solver

# set paths
THIS_DIR = Path(r'C:\Users\cjsis\Documents\Github\research\ExcelPy_Optimizer')
DATA_DIR = THIS_DIR / 'data'

# %% [markdown]
# ## Create instance of `Excel_Solver` class
# 
# This sets the following attributes:
# + `xw`: link to the Excel `book`, `sheet`, and `ranges` using xlwings.
# + `x_param`: active tuning parameters.
# + `algo_param`: algorithm method and hyperparameters.

# %%
# create instance of solver
book = THIS_DIR / "optimizer_demo.xlsx"
solver = Excel_Solver(
    book=book, sheet_name="project", 
    param_rg_name="pySolve_Param", algo_rg_name="pySolve_Algo"
)

# %%
# print attributes of solver instance
print(f"book={solver.xw.book.name}")
print(f"x={solver.x_param['param']}")
print(f"method={solver.algo_param['method']}")

# %% [markdown]
# ## Run `Excel_Solver.optimize()` method
# 
# Solves optimization problem according to `x_param` and `algo_param` attributes.

# %%
# modify algorithm parameters
algo_method = [None, 'basinhopping', 'differential_evolution', 'shgo', 'dual_annealing', 'direct'][0]
if algo_method:
    opt_params = solver.get_algo_params(method=algo_method)
    opt_params['bounds'] = solver.algo_param['param']['bounds']
    solver.set_algo_params(method=algo_method, param=opt_params)

# %%
# use `optimize` method to solve
result = solver.optimize()
print(result)

# %% [markdown]
# ## Print results from optimization

# %%
# print candidate solutions to sheet
solver.print_solutions()

# %%
# write candidate solutions to sheet and evaluate results
solver.write_solution_to_solver_range(idx=5)


