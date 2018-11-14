# monte_carlo_pi
![sample image](https://raw.githubusercontent.com/bthaman/monte_carlo_pi/master/images/monte_carlo_pi_1.jpg)
Demonstration of Monte Carlo method to calculate the value of Pi using Python and Excel.
I built this for my kids as a way to illustrate basic Monte Carlo methods and the Law of Large Numbers.

Python does all the calculations with vectorized numpy array operations, which greatly improves performance over iterating in Excel (~100x).

Input/output and visualization is done in Excel, which calls the Python method using xlwings.

