# Changelog

All notable changes to this project (`ExcelPy_Optimizer`) are to be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.0] - 2025-07-20

### Added

* Makes a `src` folder, adds `pyproject` file.

### Changed

* None

### Fixed

* Handles potential COM errors that may be encountered during an optimization.
  * `_objective_function()` method now uses a "retry + delay" mechanism to avoid Excel COM errors.
  * TODO: add methods for saving and loading the state of the optimizer.

### Notes

* None

## [0.3.0] - 2024-01-16

### Added

- In `Excel_Solver` class,
  1. Added `to_file()` method for file management for saving class instances. Can also load instances from JSON.
  2. Added `copy_figure_to_solution_sheet()` method to add post-processing functionality.

### Changed

- Improved `optimizer_demo` for more clarity on usage of `Excel_Solver` class.
  1. Added `.ipynb` and `.py` files to show how to use either UI.
- For `XW` class,
  1. Moved class definition to `utils.file_excel`.
  2. Modified `utils.json` to handle the serializing and deserializing of a `XW` class instance.
- On some method signatures of `Excel_Solver`, user can now provide references to `xw.Sheet` objects so that sheets from outside `solver.xw.book` can be referenced.
- In `print_solutions()` method, more output info provided.

### Fixed

- Using a `check_file` in the `objective_function()` method as a temporary solution to the `_solver_callback` issue.

### Notes

1. TODO: modify `kwargs` handling in `optimize()` method.
2. `to_file()` method works for json files; for the `pickle` files, need to improve `save` method to handle `XW` class instance and also the `load` method.

## [0.2.0] - 2024-01-03

## Added

- Created changelog file.
- Added `write_solution_to_solver_range(idx)` method which allows the user to write candidate solution `idx` to the Excel sheet and then evaluate the result. This is useful for `ipynb` setups.

### Changed

- Cleaner attributes for `Excel_Solver` class and improved docstring for `init` method.
  1. Added `XW` class as a convenience class to separate `xlwings` attributes from other attributes.
  2. All attributes are grouped into dictionaries for improved readability.
- More output info provided from `print_solutions()` method.
- Renamed `demo.xlsx` to `optimizer_demo.xlsx` and added features to workbook.

### Notes

1. `_solver_callback()` method still not working!
2. TODO: modify `kwargs` handling in `optimize()` method.

## [0.1.0] - 2023-12-17

### Added

- First version (predates creation of changelog file)
- Added `Excel_Solver` class for handling the mechanics of the optimization process.
  1. Interaction with Excel accomplished with `xlwings` package. Links established in `__init__` method.
  2. Optimization algorithms are from `scipy.optimize` package.
  3. `optimize()` is method for performing the optimization: working.
  4. `print_solutions()` is method for writing the solutions to an Excel sheet: working.
- Added `kwargs` for all available optimizer methods in `scipy.optimize`.
- Added `demo.xlsx` file to demonstrate use of the `Excel_Solver` class.

### Notes

1. `_solver_callback()` method defined but not working!
