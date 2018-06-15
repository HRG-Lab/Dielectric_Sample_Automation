import matplotlib.pyplot as plt
import numpy as np
import pandas as pd




###### Read in .csv file
###### Strip header from csv file
test_dataframe = pd.read_csv("Y:\\ebloom\\Oil_Tests\\SOS.csv",skiprows=6,usecols=range(3))


df = pd.DataFrame(test_dataframe)


#stats = df.describe()
#print(stats)

avg = mean(df)
print(avg)


print(test_dataframe.head())

###### Plot complex permittivity

test_dataframe.plot.scatter(x='frequency',y='e\'')
plt.show()