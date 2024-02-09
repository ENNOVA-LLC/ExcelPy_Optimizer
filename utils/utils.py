"""
"""
# Define functions to `solve` optimization and `load` optimization results
def make_list_range(ranges:list[list[float]])->list:
    """Returns list from list of range lists."""
    result = []
    for start, end in ranges:
        result += list(range(start, end + 1))
    return result

def get_solution_indices(f_list:list[float], solution_tol=0.01) -> list:
    """Returns list of indices where `f_list` is less than `solution_tol`."""
    return [i for i, f in enumerate(f_list) if f < solution_tol]

if __name__ == "__main__":
    print("utils.py module is to be imported!")