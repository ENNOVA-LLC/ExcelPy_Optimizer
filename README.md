# Excel Solver for Optimization

This folder contains a Python script, `optimizer.py`, which uses the  library to optimize parameters specified in a given Excel range named `pySolve_Param`. This script is designed to interface with Excel using , allowing for seamless integration of Python's powerful optimization capabilities with Excel's user-friendly interface.

## Features

- Optimization of parameters directly from an Excel workbook.
- Support for various optimization algorithms from , including `minimize`, `basinhopping`, `differential_evolution`, `brute`, `shgo`, `dual_annealing`, and `direct`.
- Ability to specify optimization parameters and settings within the Excel workbook.
- Storage and display of candidate solutions meeting specific criteria.

<!--REFS-->

## Requirements

To run this script, you need:

- Python installed on your system.
- `scipy` and `xlwings` libraries installed in your Python environment.
  - Use the `environment.yml` file and follow the instructions provided in the [install.md](docs/install.md) file to create a new conda environment.
- An Excel workbook setup with `pySolve_Param` and `pySolve_Algo` ranges according to the required format.

## Usage

1. Open the `optimizer.py` script in a Python editor or IDE.
2. Modify the path to the Excel workbook in the script to point to your file.
3. Run the script. The optimization results will be written to the Excel file, and candidate solutions will be printed in a new sheet named "Solutions".

Refer to the `optimizer_demo.xlsx` file for an example of how to configure the optimization.

## Class Overview

The `Excel_Solver` class in `optimizer.py` handles the optimization process. Key methods include:

- `optimize()`: Runs the optimization algorithm based on parameters and settings defined in the Excel workbook.
- `print_solutions(sheet_name)`: Writes the candidate solutions to a new `sheet_name` in the Excel workbook.
- `write_solution_to_solver_range(idx)`: Write the candidate solution `idx` to the Excel solver range.
- `close_excel`: Closes the Excel workbook and saves changes.

The script uses ranges named `pySolve_Param` and `pySolve_Algo` in the Excel workbook to read optimization variables and settings, respectively.

[scipy_optimize]: https://docs.scipy.org/doc/scipy/reference/optimize.html#optimization
[xlwings]: https://docs.xlwings.org/en/stable/quickstart.html
