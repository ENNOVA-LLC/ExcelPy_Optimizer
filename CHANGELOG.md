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

## 0.3.0

* **Date**: 2024-01-16
* **Major** revisions:
  1. In `Excel_Solver` class,  
     1. Added `to_file()` method for file management for saving class instances. Can also load instances from JSON.
        1. Works with `json` files, For the `pickle` files, need to improve `save` method to handle `XW` class instance and also the `load` method.
     2. Added `copy_figure_to_solution_sheet()` method to add post-processing functionality.
     3. More output info provided from `print_solutions()` method.
     4. User can now provide references to `xw.Sheet` objects so that sheets from outside `solver.xw.book` can be referenced.
  2. Improved `optimizer_demo` for more clarity on usage of `Excel_Solver` class.
     1. Added `.ipynb` and `.py` files to show how to use either UI.
  3. Using a `check_file` in the `objective_function` as a temporary solution to the `_solver_callback` issue.
  4. For `XW` class,  
     1. Moved class definition to `utils.file_excel`.
     2. Modified `utils.json` to handle the serializing and deserializing of a `XW` class instance.
* **Notes**:
  1. TODO: modify `kwargs` handling in `optimize()` method.
