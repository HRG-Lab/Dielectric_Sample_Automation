import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import OrderedDict
from scipy.optimize import curve_fit
# imports the csv file our data is on
percentage = 10
filepath = ""
A_list, B_list, materials, sample = [],[],[],[]
for percentage in np.arange(10,100,10):
     filepath = "Acetone_WD_40\\{0}acetone.csv".format(percentage)

     df = pd.read_csv(filepath)

     # singles out the column we want to use in our data
     frequency_list = df['frequency']
     # print(frequency_list.head)

     e_prime_value = df['e\'']
     # print(e_prime_value.head)

     #sets the x and y values as those values above
     x = frequency_list
     #
     y = e_prime_value
     #
     #least squares function
     #
     def f(x, A, B): # this is your 'straight line' y=f(x)
        #   print(A, B)
          return A * x + B
     #

     fit = curve_fit(f, x, y)[0] # your data x, y to fit
     print(fit)

     from scipy import stats

     #gradient, intercept, r_value, p_value, std_err = stats.linregress(x,y)
     #print("Intercept:", intercept)

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
     #
     def f(a, A, B):  # this is your 'straight line' y=f(x)
          return A * a + B
     #
     fit = curve_fit(f, x, y)[0]  # your data x, y to fit
     print(fit)


     A_list.append(fit[0])
     B_list.append(fit[1])
     materials.append('{0}acetone'.format(percentage))
     sample.append('PeR.csv')


     label = " {}Acetone".format(percentage)
     #graphs 10-90 percent data
     p = plt.plot(frequency_list, e_prime_value, '.',label = label+' Measured')[0]
     color = p.get_color()
     #places legend in the upper right corner
     plt.legend(bbox_to_anchor=(1, 1), borderaxespad=0.)
     #graphs all 10-90 percent regression lines
     plt.plot(x, fit[0]*x+fit[1],'-',color = color, label = label+' Fit')
     #plots example data
     plt.plot(frequency_list_ex, e_prime_value_ex, '--')
     plt.ylabel('e\'')
     plt.xlabel('Frequency')

#Stores the data in a csv file to be read in Inverse_Function
dictionary = OrderedDict([("Material",materials), ("A",A_list), ("B",B_list), ("Sample",sample)])
dataframe = pd.DataFrame(dictionary)
dataframe.to_csv("Acetone_Fits.csv")

plt.show()