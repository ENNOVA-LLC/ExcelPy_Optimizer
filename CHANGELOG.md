# Changelog

## 0.1.0

* **Date**: 2023-12-17
* **Major** revisions:
  1. First version (predates creation of changelog file)
  2. Created `Excel_Solver` class for handling the mechanics of the optimization process.
     1. Interaction with Excel accomplished with `xlwings` package.
     2. Optimization algorithms are from `scipy.optimize` package.
  3. `optimize()` is method for performing the optimization: working.
  4. `print_solutions()` is method for writing the solutions to an Excel sheet: working.
  5. Added `kwargs` for all available optimizer methods in `scipy.optimize`.
  6. Created `demo.xlsx` file to demonstrate use of the `Excel_Solver` class.
* **Notes**:
  1. `_solver_callback()` method defined but not working!

## 0.2.0

* **Date**: 2024-01-03
* **Major** revisions:
  1. Created changelog file.
  2. Cleaner attributes for `Excel_Solver` class and improved docstring for `init` method.
     1. Added `XW` class as a convenience class to separate `xlwings` attributes from other attributes.
     2. All attributes are grouped into dictionaries for improved readability.
  3. More output info provided from `print_solutions()` method.
  4. Added `write_solution_to_solver_range(idx)` method which allows the user to write candidate solution `idx` to the Excel sheet and then evaluate the result. This is useful for `ipynb` setups.
  5. Renamed `demo.xlsx` to `optimizer_demo.xlsx` and added features to workbook.
* **Notes**:
  1. `_solver_callback()` method still not working!
  2. TODO: modify `kwargs` handling in `optimize()` method.
