# Optimization Environment Setup

This guide will walk you through setting up a dedicated Conda environment for the optimization package.

## Step 1: Install Conda

If Conda is not already installed on your system, download and install [Miniconda](https://docs.conda.io/en/latest/miniconda.html) or [Anaconda](https://www.anaconda.com/products/distribution).

## Step 2: Create an `environment.yml` File

Create an `environment.yml` file in your project directory with the following content:

```yaml
name: optimization-env
channels:
  - defaults
dependencies:
  - python=3.11  # or any version you prefer
  - numpy
  - scipy
  - pandas
  - xlwings
  - matplotlib  # Optional, for plotting
  - jupyter     # Optional, for interactive notebooks
  # Add any additional packages you need
```

This file is contained in the root directory of the project.

## Step 3: Create the Conda Environment

Navigate to the project directory in the terminal or Anaconda Prompt and run:

```bash
conda env create -f environment.yml
```

## Step 4: Activate the Environment

Activate the new environment:

```bash
conda activate optimization-env
```

To check the packages installed in the active environment, run the following command:

```bash
conda list
```

## Step 5: Deactivate the Environment

When you are finished working, you can deactivate the environment:

```bash
conda deactivate
```

## Step 6: Managing Dependencies

- **To add a new package**: `conda install package-name` or update `environment.yml` and run `conda env update -f environment.yml`.
- **To update a package**: `conda update package-name`.
- **To remove a package**: `conda remove package-name`.

## Step 7: Using the Environment in VSCode

- Open VSCode.
- Open the Command Palette (`Ctrl+Shift+P` or `Cmd+Shift+P` on Mac).
- Type `Python: Select Interpreter`.
- Choose the interpreter from your Conda environment.

## Step 8: Exporting Your Environment

To share or recreate your environment:

```bash
conda env export > environment.yml
```

## Step 9: Removing the Environment

If you no longer need the environment:

```bash
conda env remove -n optimization-env
```

Remember to activate your environment each time you work on this project to ensure you are using the correct dependencies.
