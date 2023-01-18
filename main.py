# -*- coding: utf-8 -*-
"""
Created on Tue Oct 11 11:46:40 2022

@author: Jon Olav Båtbukt
"""

# - - - Setup - - -
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pyomo.opt import SolverFactory
import pyomo.environ as pyo
import time 

# - - - Data - - - 
# Trondheim - Gathered 11.oct 2022
# Oslo - Gathered 11.nov 2022
# Tromsø - Gathered 4.des 2022
# Consumption data: https://www.nordpoolgroup.com/en/Market-data1/Power-system-data/Consumption1/Consumption/NO/Hourly1/?view=table
# Price data: https://www.nordpoolgroup.com/en/Market-data1/Dayahead/Area-Prices/ALL1/Hourly/?dd=Tr.heim&view=table
demandDataNO1 = pd.read_excel("NordPoolConsNO1Sep2022.xlsx", "Data", usecols = 'B:AE') # MWh
demandDataNO1.name = "demandDataNO1"
priceDataOslo = pd.read_excel("NordPoolPriceOsloSep2022.xlsx", "Data", usecols = 'B:AE') # NOK/MWh
priceDataOslo.name = "priceDataOslo"

demandDataNO3 = pd.read_excel("NordPoolConsNO3Sep2022.xlsx", "Data", usecols = 'B:AE') # 
demandDataNO3.name = "demandDataNO3"
priceDataTrheim = pd.read_excel("NordPoolPriceTrheimSep2022.xlsx", "Data", usecols = 'B:AE') # NOK/MWh
priceDataTrheim.name = "priceDataTrheim"

demandDataNO4 = pd.read_excel("NordPoolConsNO4Sep2022.xlsx", "Data", usecols = 'B:AE') # MWh
demandDataNO4.name = "demandDataNO4"
priceDataTromso = pd.read_excel("NordPoolPriceTromsoSep2022.xlsx", "Data", usecols = 'B:AE') # NOK/MWh
priceDataTromso.name = "priceDataTromso"

# Constants
nrHours = 720 # Hours in september (30 days)
peopleInHousehold = 4 # Guesstimate on amount of people in a household
popConst_Trheim = peopleInHousehold/737 
popConst_Oslo = peopleInHousehold/2742
popConst_Tromso = peopleInHousehold/482
mva = 0.25 # 25% VAT

# Class to define each power region
class Region:
    def __init__(self, name, demand, price, popConst):
        self.name = name
        self.demand = demand
        self.price = price
        self.popConst = popConst

# Defining each power region
NO1 = Region("NO1",demandDataNO1, priceDataOslo, popConst_Oslo)
NO3 = Region("NO3",demandDataNO3, priceDataTrheim, popConst_Trheim)
NO4 = Region("NO4",demandDataNO4, priceDataTromso, popConst_Tromso)

# Adding all power regions to an array
Regions = [NO1, NO3, NO4]


# - - - Functions - - - 

# Writes all demand and prices to Excel - - - - - - - - - - - - - - -
def toExcel():
    NO1_df = pd.DataFrame()
    NO3_df = pd.DataFrame()
    NO4_df = pd.DataFrame()
    Comparison_df = pd.DataFrame()
    
    # Oslo
    P_load_NO1 = []
    p_NO1 = []
    
    # Trheim
    P_load_NO3 = []
    p_NO3 = []
    
    # Tromsø
    P_load_NO4 = []
    p_NO4 = []
    
    # Can do all areas with only iterating over NO1, because they all have the same size (720 hours = 30 days * 24 hours)
    for i in demandDataNO1.columns:
        for j in demandDataNO1.index:
            P_load_NO1.append(demandDataNO1[i].iloc[j] * popConst_Oslo)
            p_NO1.append(priceDataOslo[i].iloc[j] * 10**-3)
            P_load_NO3.append(demandDataNO3[i].iloc[j] * popConst_Trheim)
            p_NO3.append(priceDataTrheim[i].iloc[j] *10**-3)
            P_load_NO4.append(demandDataNO4[i].iloc[j] * popConst_Tromso)
            p_NO4.append(priceDataTromso[i].iloc[j] *10**-3)
    
    NO1_df["Demand [kWh]"] = P_load_NO1
    NO1_df["Price [NOK/kWh]"] = p_NO1
    NO3_df["Demand [kWh]"] = P_load_NO3
    NO3_df["Price [NOK/kWh]"] = p_NO3
    NO4_df["Demand [kWh]"] = P_load_NO4
    NO4_df["Price [NOK/kWh]"] = p_NO4
    
    # Comparison
    Comparison_df["NO1 [kWh]"] = P_load_NO1
    Comparison_df["NO1 [NOK/kWh]"] = p_NO1
    Comparison_df["NO3 [kWh]"] = P_load_NO3
    Comparison_df["NO3 [NOK/kWh]"] = p_NO3
    Comparison_df["NO4 [kWh]"] = P_load_NO4
    Comparison_df["NO4 [NOK/kWh]"] = p_NO4
    
    with pd.ExcelWriter("Demand and prices.xlsx", engine='openpyxl', mode = 'a', if_sheet_exists='overlay') as writer:
        NO1_df.to_excel(writer, sheet_name  = "NO1")
        NO3_df.to_excel(writer, sheet_name  = "NO3")
        NO4_df.to_excel(writer, sheet_name  = "NO4")
        Comparison_df.to_excel(writer, sheet_name  = "Comparison")
        
    return


# Calculates baseline cost - no battery - - - - - - - - - - - - - - -
def baselineCost(NOx, writeOut):
    cost = 0
    
    for i in NOx.demand.columns: 
        for j in NOx.demand.index: 
            cost += float(NOx.demand[i].iloc[j]) * NOx.popConst * float(NOx.price[i].iloc[j]) * 10**-3 # kWh / person

    if(NOx.name != "NO4"):
        cost = cost * (1 + mva)

    format_cost = "{:.2f}".format(cost)
    if(writeOut):
        print("No battery, money spent:", format_cost, "NOK")
    
    return cost


# Calculate cost  with battery - - - - - - - - - - - - - - -
def batteryCost(NOx, writeOut):
    
    # df and list for gathering battery data in Excel    
    tb_df = pd.DataFrame()
    tb_P_imp = []
    tb_B = []
    tb_P_ch = []
    tb_P_dis = [] 

    # Create model
    model = pyo.ConcreteModel() 
    
    # Parametres (demand + price)
    P_load = []
    p = []
    for i in NOx.demand.columns: 
       for j in NOx.demand.index: 
            P_load.append(NOx.demand[i].iloc[j] * NOx.popConst)
            p.append(NOx.price[i].iloc[j] * 10**-3)
            
    n_ch = 0.92 # Charging efficiency
    n_dis = 0.95 # Discharging efficiency
    B_f = 5 # Battery level in the first hour
    B_l = B_f # Battery level in the last hour
    B_capMax = 9 # Battery max capacity
    B_chMax = 4.5 # Max charging rate
    B_disMax = 4.5 # Max discharge rate
    vatVar = 1 # Base VAT-multiplier 

    # Creates Variables: 
    model.P_imp = pyo.Var(range(nrHours), within = pyo.NonNegativeReals) # Power imported from grid
    model.P_ch = pyo.Var(range(nrHours), within = pyo.NonNegativeReals)  # Power charged to battery
    model.P_dis = pyo.Var(range(nrHours), within = pyo.NonNegativeReals)  # Power discharged from battery
    model.B = pyo.Var(range(nrHours + 1), within = pyo.NonNegativeReals) # Battery level. Has [+1] so the last hour of the month equals the f
    
    if(NOx.name != "NO4"):
        vatVar += mva
    
    # Creates objective function: 
    obj = sum(model.P_imp[i] * p[i] * vatVar for i in range(nrHours))
    model.objFunc = pyo.Objective(expr = obj, sense = pyo.minimize)
    
    # Create constraints: 
    constraints = []
    constraints.append(model.B[0] == B_f)

    for i in range(nrHours):
        constraints.append(model.P_imp[i] + (model.P_dis[i] - model.P_ch[i]) == P_load[i]) #Power balance
        constraints.append(model.P_ch[i] <= B_chMax) # Max charge rate
        constraints.append(model.P_dis[i] <= B_disMax) # Max discharge rate
        constraints.append(model.B[i] <= B_capMax) # Capacity of battery!
        constraints.append(model.B[i] + model.P_ch[i]*n_ch - model.P_dis[i]/n_dis == model.B[i + 1]) # Battery balance    
    
    # Constraints for last hour
    constraints.append(model.B[720] == B_l) # Must start the next month with same battery level as the previous month
    
    model.constraints = pyo.ConstraintList()
    for constraint in constraints:
        model.constraints.add(expr = constraint)
        
    opt = SolverFactory("gurobi", solver_io="python")
    
    results = opt.solve(model)
    
    # Prepare adding data to Excel
    for i in range(nrHours):
        tb_B.append(model.B.get_values()[i])
        tb_P_ch.append(model.P_ch.get_values()[i])
        tb_P_dis.append(model.P_dis.get_values()[i])
        tb_P_imp.append(model.P_imp.get_values()[i])
        
        # Check if battery is charging and discharging at the same time
        if(model.P_ch.get_values()[i] > 0 and model.P_dis.get_values()[i] > 0):
            print("Hour", i, ": Charge and discharge at the same time!")
        
    
    if(writeOut):
        format_cost = "{:.2f}".format(model.objFunc())
        print("All-knowing battery, money spent:", format_cost, "NOK")
        
    tb_df["Power imported"] = tb_P_imp
    tb_df["Battery charged"] = tb_P_ch
    tb_df["Battery discharged"] = tb_P_dis
    tb_df["Battery level"] = tb_B
    tb_df["Baseline cost"] = baselineCost(NOx, 0)
    tb_df["Cost with battery"] = model.objFunc()

    with pd.ExcelWriter("Test battery data.xlsx", engine='openpyxl', mode = 'a', if_sheet_exists='overlay') as writer:
        tb_df.to_excel(writer, sheet_name  = NOx.name)
    
    
    return model.objFunc()


# Compares the baseline case with the battery case - - - - - - - - - - - - - - -
# If writeOut = True, then the cost for each case is also printed, not just the savings
def compareBaselineToBattery(writeOut):
    
    startTime = time.time()
    
    for NOx in Regions:
        print("- - - Region: ", NOx.name, " - - -")
        costNormal = baselineCost(NOx,writeOut)
        costBattery = batteryCost(NOx,writeOut)

        diff_cost = "{:.2f}".format(costNormal - costBattery)
        print("Money saved:", diff_cost, "NOK \n")
        
    executionTime = "{:.4f}".format(time.time() - startTime)
    
    if(writeOut):
        print("Total calculation time:", executionTime, "sec")

    return


# Creates heatmap for demand - - - - - - - - - - - - - - -
def demandHeatmap(NOx):
    demand = np.array(NOx.demand)
    
    fig, ax = plt.subplots()
    im = ax.imshow(demand, cmap = "plasma")
    cbar = ax.figure.colorbar(im, 
                          ax = ax,
                          shrink=0.5)
    cbar.ax.set_ylabel("MWh", rotation=-90, va="bottom")
    
    days = ["1","2","3","4","5","6","7","8","9","10",
             "11","12","13","14","15","16","17","18",
             "19","20","21","22","23","24","25","26",
             "27","28","29","30"]

    ax.set_xticks(np.arange(len(days)), labels = days)
    ax.set_yticks(np.arange(24))
    
    # Removes every other x-ticket label, so it's less cluttered
    for label in ax.xaxis.get_ticklabels()[::2]:
        label.set_visible(False)
    
    # Removes every other y-ticket label, so it's less cluttered
    for label in ax.yaxis.get_ticklabels()[::2]:
        label.set_visible(False)
    
    plt.setp(ax.get_xticklabels(), rotation=40, ha="right", rotation_mode="anchor")
    plt.xlabel("Days")
    plt.ylabel("Hours")
    
    if(NOx.name == "NO1"):
        ax.set_title("Demand for NO1")
    elif (NOx.name == "NO3"):
        ax.set_title("Demand for NO3")
    else:
        ax.set_title("Demand for NO4")

    fig.tight_layout()
    plt.show()
        
    return 


# Creates heatmap for price - - - - - - - - - - - - - - -
def priceHeatmap(NOx):
    price = np.array(NOx.price)
    
    fig, ax = plt.subplots()
    im = ax.imshow(price, cmap = "plasma")
    cbar = ax.figure.colorbar(im, 
                          ax = ax,
                          shrink=0.5)
    cbar.ax.set_ylabel("NOK/MWh", rotation=-90, va="bottom")
    
    days = ["1","2","3","4","5","6","7","8","9","10",
         "11","12","13","14","15","16","17","18",
         "19","20","21","22","23","24","25","26",
         "27","28","29","30"]

    ax.set_xticks(np.arange(len(days)), labels = days)
    ax.set_yticks(np.arange(24))
    
    # Removes every other x-ticket label, so it's less cluttered
    for label in ax.xaxis.get_ticklabels()[::2]:
        label.set_visible(False)
    
    # Removes every other y-ticket label, so it's less cluttered
    for label in ax.yaxis.get_ticklabels()[::2]:
        label.set_visible(False)
    
    plt.setp(ax.get_xticklabels(), rotation=40, ha="right", rotation_mode="anchor")
    plt.xlabel("Days")
    plt.ylabel("Hours")
    
    if(NOx.name == "NO1"):
        ax.set_title("Price for Oslo")
    elif (NOx.name == "NO3"):
        ax.set_title("Price for Trondheim")
    else:
        ax.set_title("Price for Tromsø")

    fig.tight_layout()
    plt.show()
        
    return 


# Creates heatmaps for both demand and price for all regions  - - - - - - - - - - - - - - -
def createHeatmaps():
    for NOx in Regions:
        demandHeatmap(NOx)
        priceHeatmap(NOx)
        
    return
    


# - - - Main - - -
compareBaselineToBattery(True)
#createHeatmaps()
toExcel()

