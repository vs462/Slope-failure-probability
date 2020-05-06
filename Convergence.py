import pandas as pd
import xlwings as xw
import numpy as np
import time

start_time = time.time()
"""
Open Template (window will stay open during calculation)
"""

wb = xw.Book('MonteCarlo.xlsx')

sht = wb.sheets("Input and results")
# Get input page open

N_conv=10 #number of simulations for convergence 
N_summ=1 #number of simulations for results analysis 

ConvergenceLists =[]
SummaryLists =[]

for i in range(N_conv): ConvergenceLists.append([i]) #create list of lists    
for i in range(N_summ): SummaryLists.append([i]) #create list of lists    


for item in ConvergenceLists:
    #power = random.randint(0,4)
    #N = np.power(10, power)
    print (item)
    power = np.random.uniform(0,1)
    N = int(np.power(10000, power))
    sht.range('E14').value = N #test fo N being randomly selected from 0 to 1000000
    results = sht.range('G19:G21').value  
    item.append(N)
    item.extend(results)

    
for item in SummaryLists:
    sht.range('E14').value = 10000 #to make the model reset the random variables
    results = sht.range('G19:G25').value    
    item.extend(results)

#turn into tables 
dfConv = pd.DataFrame(ConvergenceLists)
dfSumm = pd.DataFrame(SummaryLists)

print (dfConv)

# print results to csv
dfConv.to_csv('convergence3.csv',index=False, header=["num","N", "Probability of failure", "NA", "Average FoS" ])
dfSumm.to_csv('summary3.csv',index=False, header=["num", "Probability of failure", "NA", "Average FoS", "Standard Deviation", "Max", "Min", "Skewness" ])

print("--- %s seconds ---" % (time.time() - start_time))
