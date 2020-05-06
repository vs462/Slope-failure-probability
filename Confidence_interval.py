import pandas as pd
import xlwings as xw
import random
import numpy as np
import time

start_time = time.time()


target = 10000
beta = 1.0/target

Y = np.random.exponential(10000)
print(Y)


"""
Open Template (window will stay open during calculation)
"""


wb = xw.Book('MonteCarlo.xlsx')

sht = wb.sheets("Input and results")
# Get input page open

N=50 #number of simulations
ListOfLists =[]

for i in range(N): ListOfLists.append([i]) #create list of lists    


for item in ListOfLists:
    N = random.randint(0,10000)
    sht.range('E14').value = N #test fo N being randomly selected from 0 to 1000000
    results = sht.range('G19:G25').value    

    item.extend(results)
    item.append(N)
    
print (ListOfLists)    


df = pd.DataFrame(ListOfLists)
print (df)

# print results to csv
df.to_csv('results.csv',index=False, header=["num","N", "Probability of failure", "NA", "Average FoS", "Standart Deviation of FoS", "Minimum FoS", "Maximum FoS", "Skewness of FoS"])

print("--- %s seconds ---" % (time.time() - start_time))         
