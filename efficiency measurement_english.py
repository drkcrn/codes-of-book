from gurobipy import *
import xlrd
import xlwt
from collections import defaultdict

ask1 = input("How many DMUs are in your data set?")
n = int(ask1)
ask2 = input("How many inputs are in your data set?")
m = int(ask2)
ask3 = input("How many outputs are in your data set?")
s = int(ask3)
ask4 = input("Which data file would you like to work with?")
file_name = ask4 + ".xls"
book = xlrd.open_workbook(file_name)
ask5 = input("Which sheet would you like to work with?")
sheet = book.sheet_by_index(int(ask5))

x = [[sheet.cell_value(r,c) for r in range(1,sheet.nrows)] for c in range(1, m+1)]
y = [[sheet.cell_value(r,c) for r in range(1,sheet.nrows)] for c in range(m+1, m+s+1)]

virtualx,virtualy = {},{}

# CRS_IO (input-oriented CRS multiplier model)
crs_io = {}
virtual_x = defaultdict(dict)
virtual_y = defaultdict(dict) 
for o in range(n):
    M = Model()
    v,μ = {},{}
    for i in range(m):
        v[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")  
    for r in range(s):
        μ[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MAXIMIZE
    M.update()
    M.setObjective(sum(μ[r]*y[r][o] for r in range(s)))
    for j in range(n):
        M.addConstr(sum(μ[r]*y[r][j] for r in range(s)) <= sum(v[i]*x[i][j] for i in range(m)))
    M.addConstr(sum(v[i]*x[i][o] for i in range(m)) == 1)
    M.optimize()
    crs_io[o] = M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["CRS_IO"] = virtual_x
virtualy["CRS_IO"] = virtual_y 

# CRS_OO (output-oriented CRS multiplier model)
crs_oo = {}
virtual_x = defaultdict(dict)
virtual_y = defaultdict(dict) 
for o in range(n):
    M = Model()
    v,μ = {},{}
    for i in range(m):
        v[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")  
    for r in range(s):
        μ[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MINIMIZE
    M.update()
    M.setObjective(sum(v[i]*x[i][o] for i in range(m)))
    for j in range(n):
        M.addConstr(sum(v[i]*x[i][j] for i in range(m)) >= sum(μ[r]*y[r][j] for r in range(s)))
    M.addConstr(sum(μ[r]*y[r][o] for r in range(s)) == 1)
    M.optimize()
    crs_oo[o] = 1/M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["CRS_OO"] = virtual_x
virtualy["CRS_OO"] = virtual_y 

# VRS_IO (input-oriented VRS multiplier model)
vrs_io = {}
virtual_x = defaultdict(dict)
virtual_y = defaultdict(dict) 
for o in range(n):
    M = Model()
    v,u = {},{}
    for i in range(m):
        v[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")  
    for r in range(s):
        μ[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    μ0 = M.addVar(lb = -GRB.INFINITY, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MAXIMIZE
    M.update()
    M.setObjective(sum(μ[r]*y[r][o] for r in range(s)) + μ0)
    for j in range(n):
        M.addConstr(sum(μ[r]*y[r][j] for r in range(s)) + μ0 <= sum(v[i]*x[i][j] for i in range(m)))
    M.addConstr(sum(v[i]*x[i][o] for i in range(m)) == 1)
    M.optimize()
    vrs_io[o] = M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["VRS_IO"] = virtual_x
virtualy["VRS_IO"] = virtual_y 
        
# VRS_OO (output-oriented VRS multiplier model)
vrs_oo = {}
virtual_x = defaultdict(dict)
virtual_y = defaultdict(dict) 
for o in range(n):
    M = Model()
    v,u = {},{}
    for i in range(m):
        v[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")  
    for r in range(s):
        μ[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    v0 = M.addVar(lb = -GRB.INFINITY, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MINIMIZE
    M.update()
    M.setObjective(sum(v[i]*x[i][o] for i in range(m)) + v0)
    for j in range(n):
        M.addConstr(sum(v[i]*x[i][j] for i in range(m)) + v0 >= sum(μ[r]*y[r][j] for r in range(s)))
    M.addConstr(sum(μ[r]*y[r][o] for r in range(s)) == 1)
    M.optimize()
    vrs_oo[o] = 1/M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["VRS_OO"] = virtual_x
virtualy["VRS_OO"] = virtual_y

wb = xlwt.Workbook()
models = {"CRS_IO":crs_io, "CRS_OO":crs_oo, "VRS_IO":vrs_io, "VRS_OO":vrs_oo}
for model in models:
    ws = wb.add_sheet(model)
    ws.write(0,0,"Score")
    for i in range(1,m+1):
        string = str(i)
        ws.write(0,i,"Virtual Input_"+string)
    for r in range(m+1,m+s+1):
        string = str(r-m)
        ws.write(0,r,"Virtual Output"+string)
    for j in range(n):
        ws.write(j+1,0,models[model][j])
        for i in range(m):
            ws.write(j+1,i+1,virtualx[model][i][j])
        for r in range(s):
            ws.write(j+1,m+r+1,virtualy[model][r][j])
wb.save("Efficiency Measurement.xls")
