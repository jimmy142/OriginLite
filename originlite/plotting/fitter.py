from dataclasses import dataclass
from typing import Callable, List
import numpy as np
from scipy.optimize import curve_fit

@dataclass
class FitResult:
    params: np.ndarray
    covariance: np.ndarray | None
    xfit: np.ndarray
    yfit: np.ndarray

def fit_and_predict(x: np.ndarray,
                    y: np.ndarray,
                    model: Callable,
                    p0: List[float] | None = None,
                    bounds = (-np.inf, np.inf),
                    npts: int = 1000) -> FitResult:
    popt, pcov = curve_fit(model, x, y, p0=p0, bounds=bounds, maxfev=10000)
    xfit = np.linspace(np.nanmin(x), np.nanmax(x), npts)
    yfit = model(xfit, *popt)
    return FitResult(params=popt, covariance=pcov, xfit=xfit, yfit=yfit)
