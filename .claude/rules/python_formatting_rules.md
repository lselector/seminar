
All Python scripts in this project
must follow these rules

Each python script file should start with a multiline doc  
concisely explaining the purpose and function
of this file.

It should show example(s) of usage.

Also when it was created and last updated.

All function and class definitions should have short docstrings.
Preferably to keep them to concise one-liners.

The total length of any function must not exceed 35 lines
including function definition line and all code,
not counting docstring, comments and empty lines

If a function grows big, split it into smaller functions.

Every external function definition must be preceded 
by a horizontal separator line 64 characters long
in the following format:

```python
# --------------------------------------------------------------
```

Every internal function definition 
(function inside a class or inside another function)
must be preceded by a horizontal separator
started indented same way as the function,
and ending after position 44 on the line
(so it is usually 40 characters long).

**Format:**
```python
    # --------------------------------------
```

**Rules for separator lines:**
- Starts with `#` (pound symbol)
- Followed by one space
- Followed by exactly 62 dash characters (`-`) for external functions,
  or followed by 38 characters for internal functions

The `if __name__ == "__main__":` statement must also be preceded by a separator line.

Length of all lines of the code or comments 
should not be longer than 65 characters

When importing common modules, use the following
common standard ways of naming them:

```python
import datetime as dt
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import polars as pl
```

Keep the lengths of individual python scripts 
and modules no longer than 800 lines. 
If file grows longer, split it.

These rules apply to all Python files in the project and subdirectories
