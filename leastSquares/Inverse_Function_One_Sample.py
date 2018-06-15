import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

#Mixture we are trying to predict
dg = pd.read_csv('Crude_Oil\\PeR.csv')

#singles out the column we want to use in our data
#print(filepath)
frequency_list_ex = dg['frequency']
# print(frequency_list.head)

e_prime_value_ex = dg['e\'']
# print(e_prime_value.head)

a = frequency_list_ex
    #
b = e_prime_value_ex
#
# least squares function
from scipy.optimize import curve_fit
#
def f(a, A, B):  # this is your 'straight line' y=f(x)
     return A * a + B
#
fit = curve_fit(f, a, b)[0]  # your data x, y to fit
print(fit)

#label = " {}Acetone".format(percentage)
#graphs 10-90 percent data
#p = plt.plot(frequency_list, e_prime_value, '.',label = label+' Measured')[0]
#color = p.get_color()
#places legend in the upper right corner
plt.legend(bbox_to_anchor=(1, 1), borderaxespad=0.)
#graphs all 10-90 percent regression lines
#plt.plot(a, fit[0]*a+fit[1],'-')
#plots example data
plt.plot(frequency_list_ex, e_prime_value_ex, '--')
plt.ylabel('e\'')
plt.xlabel('Frequency')

plt.show()