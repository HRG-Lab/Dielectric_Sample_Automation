import pandas as pd
#import numpy as np
import matplotlib.pyplot as plt

# imports the csv file our data is on
#percentage = 10
#filepath = ""
#for percentage in np.arange(10,100,10):
      #filepath = "Acetone_WD_40\\acetone.csv"
df = pd.read_csv('Crude_Oil\\PeR.csv')

     #singles out the column we want to use in our data
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

#plots all the data into a graph
plt.plot(x, y)
plt.axis([1e9, 6e9, 3, 5])
plt.ylabel('e\'')
plt.xlabel('Frequency')
plt.title('Permian Basin')

plt.show()