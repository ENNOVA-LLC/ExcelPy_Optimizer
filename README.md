# Excel Solver for Optimization

This folder contains a Python module - `optimizer.py` - that contains the `Excel_Solver` class definition. The `optimizer.py` file sits in the `\ROOT` folder.

`Excel_Solver` uses the [`scipy.optimize`](https://docs.scipy.org/doc/scipy/reference/optimize.html#optimization) library to optimize parameters specified in a given Excel named range `pySolve_Param`. This script is designed to interface with Excel using [`xlwings`](https://docs.xlwings.org/en/stable/quickstart.html), allowing for seamless integration of Python's powerful optimization capabilities with Excel's user-friendly interface.

## Features

- Optimization of parameters directly from an Excel workbook.
- Support for various optimization algorithms from `scipy.optimize`, including:
  - `minimize`, `basinhopping`, `differential_evolution`, `brute`, `shgo`, `dual_annealing`, and `direct`.
- Ability to specify optimization parameters `x` and algorithm hyperparameters `algo` within the Excel workbook.
- Storage and display of candidate solutions meeting specific `solution_tol` criteria.

## Requirements

To run this script, you need:

- Python installed on your system.
- `scipy` and `xlwings` libraries installed in your Python environment.
  - Use the `environment.yml` file and follow the instructions provided in the [install_env.md](docs/install_env.md) file to create a new conda environment for this package.
- An Excel workbook setup with `pySolve_Param` and `pySolve_Algo` ranges according to the required format.

## Usage

1. Create a copy of `optimizer_demo.py` or `optimizer_demo.ipynb` and open the file in a Python editor or IDE.
2. Modify the path to the Excel workbook in the script to point to your Excel file.
   - The `Excel_Solver` class reads the objective function value `f` from the `pySolve_Param` range, it does not calculate `f`.
   - The calculation of `f` must be performed on the sheet and update every time `x` changes.
3. Run the script. The optimization results can be written to the Excel file using the `print_solutions()` method.
   - To terminate the `Excel_Solver.optimize()` method, move the file [`stop_optimizer.txt`](docs/stop_optimizer.txt) to the same folder as the `optimize.py` module.
   - `stop_optimizer` should be stored in the `docs` folder and only moved to `\ROOT` temporarily to terminate a running optimization.

Refer to the `optimizer_demo.xlsx` file for an example of how to configure the optimization, including the expected format of the named ranges `pySolve_Param, pySolve_Algo`.

## Class Overview

The `Excel_Solver` class in the `optimizer.py` module handles the optimization process.

Key *attributes* include:

- `xw`: xlwings references to the Excel `book`, `sheet` and `ranges`.
- `x_param:dict`: The decision variables to be tuned; includes keys `{'param', 'min', 'max', 'val', 'obj'}`.
- `algo_param:dict`: The algorithm parameters that govern the chosen `scipy.optimize` method.
- `solutions:dict`: Stores the candidate solutions; includes keys `{'result', 'nfev', 'storage_tol', 'f', 'error', 'x'}`
  - `nfev` counts the total calls to the objective function method.
  - if `f(x) < storage_tol` then solution set `{f, error, x}` is stored as an element in a list.

Key *methods* include:

- `optimize()`: Runs the optimization algorithm based on parameters and algorithm settings defined in the Excel workbook.
- `print_solutions(sheet_name)`: Writes the candidate solutions to a new `sheet_name` in the Excel workbook.
- `write_solution_to_solver_range(idx)`: Write the candidate solution `idx` to the `self.x_param['val']` range.
- `copy_figure_to_solution_sheet(idx_list)`: Write the candidate solutions corresponding to the `idx_list`, copy figure from Excel sheet, and paste to `solution` sheet.
- `to_file(file_path)`: Saves the instance to a `json` or `pkl` file.

The `Excel_Solver` class uses ranges named `pySolve_Param` and `pySolve_Algo` in the Excel workbook to read optimization variables and algorithm settings, respectively.
