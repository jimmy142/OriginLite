import numpy as np
from scipy.special import wofz

def linear(x, m, c):
    return m * x + c

def exponential(x, A, k, c):
    return A * np.exp(k * x) + c

def gaussian(x, A, x0, sigma, c):
    return A * np.exp(-0.5 * ((x - x0) / sigma) ** 2) + c

def lorentzian(x, A, x0, gamma, c):
    return A * (gamma**2 / ((x - x0) ** 2 + gamma**2)) + c

def voigt(x, A, x0, sigma, gamma, c):
    z = ((x - x0) + 1j * gamma) / (sigma * np.sqrt(2))
    return A * np.real(wofz(z)) / (sigma * np.sqrt(2 * np.pi)) + c

def guess_gaussian(x, y):
    A0 = float(np.nanmax(y) - np.nanmin(y))
    x0 = float(x[np.nanargmax(y)])
    sigma0 = float((np.nanmax(x) - np.nanmin(x)) / 10)
    c0 = float(np.nanmin(y))
    return [A0, x0, sigma0, c0]

def guess_lorentzian(x, y):
    A0 = float(np.nanmax(y) - np.nanmin(y))
    x0 = float(x[np.nanargmax(y)])
    gamma0 = float((np.nanmax(x) - np.nanmin(x)) / 20)
    c0 = float(np.nanmin(y))
    return [A0, x0, gamma0, c0]

def guess_voigt(x, y):
    A0 = float(np.nanmax(y) - np.nanmin(y))
    x0 = float(x[np.nanargmax(y)])
    sigma0 = float((np.nanmax(x) - np.nanmin(x)) / 20)
    gamma0 = sigma0
    c0 = float(np.nanmin(y))
    return [A0, x0, sigma0, gamma0, c0]

def guess_exponential(x, y):
    return [float(np.nanmax(y) - np.nanmin(y)), 0.1, float(np.nanmin(y))]
