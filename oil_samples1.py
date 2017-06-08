#myFile = open("t11.prn", "r")

#print("readin...\n" + myFile.read())

#Mean = ("myFile")

import csv
from collections import Counter

def average_column (csv_filepath):
    column_totals = Counter()
    with open(csv_filepath,"rb") as f:
        reader = csv.reader(f)
        row_count = 0.0
        for row in reader:
            for column_idx, column_value in enumerate(row):
                try:
                    n = float(column_value)
                    column_totals[column_idx] += n
                except ValueError:
                    print"Error -- ({}) Column({}) could not be converted to float!".format(column_value, column_idx)
            row_count += 1.0

    row_count -= 1.0

    column_indexes = column_totals.keys()
    column_indexes.sort()

    averages = [column_totals[idx]/row_count for idx in column_indexes]
    print"The averages of e' and e'' respectively are:", averages
    return averages

average_column("t11.csv")