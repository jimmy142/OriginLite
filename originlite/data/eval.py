from __future__ import annotations
import math
import numpy as np

# Safe evaluation environment for column expressions
_SAFE_GLOBALS = {
    'np': np,
    # common math funcs
    'sin': np.sin, 'cos': np.cos, 'tan': np.tan,
    'exp': np.exp, 'log': np.log, 'log10': np.log10,
    'sqrt': np.sqrt, 'abs': np.abs, 'arctan': np.arctan,
    'sinh': np.sinh, 'cosh': np.cosh, 'tanh': np.tanh,
    'pi': math.pi, 'e': math.e,
    'min': np.minimum, 'max': np.maximum,
    'where': np.where,
}

def eval_expression(expr: str, local_symbols: dict[str, np.ndarray]) -> np.ndarray:
    """Evaluate a numpy expression using column names in local_symbols.
    Example: expr = '2*I - 0.5*V' where local_symbols={'I': array, 'V': array}
    """
    # Disallow dunder access
    if "__" in expr:
        raise ValueError("Invalid expression")
    return eval(expr, _SAFE_GLOBALS, local_symbols)  # noqa: S307 (controlled env)
