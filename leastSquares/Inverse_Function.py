import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# imports the csv file our data is on
def permittivity_from_frequency(material, frequency):
     df = pd.read_csv('Acetone_Fits.csv')
     # print(df.head)
     grouped = df.groupby('Material')
     fit = grouped.get_group(material)
     A = fit['A'].iloc[0]
     B = fit['B'].iloc[0]
     return A, B #*frequency+B
print(permittivity_from_frequency('20acetone',1e9))
