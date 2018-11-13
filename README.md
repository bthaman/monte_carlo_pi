# monte_carlo_pi
Demonstration of Monte Carlo method to calculate the value of Pi using Python and Excel.
This is a great way to illustrate basic Monte Carlo methods and the Law of Large Numbers.

Python does all the calculations using numpy arrays, which greatly improves performance over iterating in Excel (~100x).
Vectorized functions operate on the numpy arrays for speed.

Input/output and visualization is done in Excel, which calls the Python method using xlwings.

