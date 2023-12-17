# Excel Solver for Optimization

This folder contains a Python script, `optimizer.py`, which uses the `scipy.optimize` library to optimize parameters specified in a given Excel range named `pySolve_Knobs`. This script is designed to interface with Excel using `xlwings`, allowing for a seamless integration of Python's powerful optimization capabilities with Excel's user-friendly interface.

## Features

- Optimization of parameters directly from an Excel workbook.
- Support for various optimization algorithms from `scipy.optimize`, including `minimize`, `basinhopping`, `differential_evolution`, `brute`, `shgo`, and `dual_annealing`.
- Ability to specify optimization parameters and settings within the Excel workbook.
- Storage and display of candidate solutions meeting specific criteria.

## Requirements

To run this script, you need:

- Python installed on your system.
- `scipy` and `xlwings` libraries installed in your Python environment.
- An Excel workbook set up with `pySolve_Knobs` and `pySolve_Settings` ranges according to the required format.

## Usage

1. Open the `optimizer.py` script in a Python editor or IDE.
2. Modify the path to the Excel workbook in the script to point to your file.
3. Run the script. The optimization results will be written back to the Excel file, and candidate solutions will be printed in a new sheet named "Solutions".

## Script Overview

The `Excel_Solver` class in `optimizer.py` handles the optimization process. Key methods include:

- `optimize`: Runs the optimization algorithm based on parameters and settings defined in the Excel workbook.
- `print_solutions`: Writes the candidate solutions to a new sheet in the Excel workbook.
- `close_excel`: Closes the Excel workbook and saves changes.

The script uses ranges named `pySolve_Param` and `pySolve_Algo` in the Excel workbook to read optimization variables and settings, respectively.
