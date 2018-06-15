import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# imports the csv file our data is on
percentage = 10
filepath = ""
for percentage in np.arange(10,100,10):
     filepath = "Acetone_WD_40\\{0}acetone.csv".format(percentage)

     df = pd.read_csv(filepath)

     #singles out the column we want to use in our data
     #print(filepath)
     frequency_list = df['frequency']
     # print(frequency_list.head)

     e_prime_value = df['e\'']
     # print(e_prime_value.head)

     #print("first element",frequency_list.iloc[0])
     # print(df.head)
     # saved_column = df.column_name
     #
     #sets the x and y values as those values above
     x = frequency_list
     #
     y = e_prime_value
     #
     #least squares function
     from scipy.optimize import curve_fit
     #
     def f(x, A, B): # this is your 'straight line' y=f(x)
          return A*x + B
     #
     fit = curve_fit(f, x, y)[0] # your data x, y to fit
     print(fit)

     #plots all the data into a graph
     #plt.subplot(211)
     label = " {}Acetone".format(percentage)
     p = plt.plot(frequency_list, e_prime_value, '.',label = label+' Measured')[0]
     color = p.get_color()
     plt.legend(bbox_to_anchor=(1, 1), borderaxespad=0.)
     plt.plot(x, fit[0]*x+fit[1],'-',color = color, label = label+' Fit')
     plt.ylabel('e\'')
     plt.xlabel('Frequency')

plt.show()