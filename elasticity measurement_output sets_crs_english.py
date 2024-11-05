from gurobipy import *
import xlrd
import xlwt

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

XA,YA,YB,XC,YC = [],[],[],[],[]
Ain  = input("Which input variables are in set A? List them (e.g. 1 if only 1, 1-2 if 1 and 2, 1-2-3 if 1,2 and 3, 1-3 if 1 and 3, etc.)")
Aout = input("Which output variables are in set A? List them (e.g. 1 if only 1, 1-2 if 1 and 2, 1-2-3 if 1,2 and 3, 1-3 if 1 and 3, etc.)")
Bout = input("Which output variables are in set B? List them (e.g. 1 if only 1, 1-2 if 1 and 2, 1-2-3 if 1,2 and 3, 1-3 if 1 and 3, etc.)")
Cin  = input("Which input variables are in set C? List them (e.g. 1 if only 1, 1-2 if 1 and 2, 1-2-3 if 1,2 and 3, 1-3 if 1 and 3, etc.) or type N for “none”")
Cout = input("Which output variables are in set C? List them (e.g. 1 if only 1, 1-2 if 1 and 2, 1-2-3 if 1,2 and 3, 1-3 if 1 and 3, etc.) or type N for “none”")
if Ain == "N":
    XA = []
else:
    l1 = [int(k) for k in Ain.split("-")]
    for k in l1:
        XA.append(x[int(k)-1])        
if Aout == "N":
    YA = []
else:
    l2 = [int(k) for k in Aout.split("-")]
    for k in l2:
        YA.append(y[int(k)-1])
if Bout == "N":
    YB = []
else:
    l3 = [int(k) for k in Bout.split("-")]
    for k in l3:
        YB.append(y[int(k)-1])
if Cin == "N":
    XC = []
else:
    l4 = [int(k) for k in Cin.split("-")]
    for k in l4:
        XC.append(x[int(k)-1])
if Cout == "N":
    YC = []
else:
    l5 = [int(k) for k in Cout.split("-")]
    for k in l5:
        YC.append(y[int(k)-1])

epsilonP = {}
for o in range(n):
    M = Model()
    vA,vC,μA,μB,μC = {},{},{},{},{}
    for i in range(len(XA)):
        vA[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for i in range(len(XC)):
        vC[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YA)):
        μA[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YB)):
        μB[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YC)):
        μC[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MINIMIZE
    M.update()
    M.setObjective(sum(vA[i]*XA[i][o] for i in range(len(XA))) - sum(μA[r]*YA[r][o] for r in range(len(YA))))
    M.addConstr(sum(vA[i]*XA[i][o] for i in range(len(XA)))
                + sum(vC[i]*XC[i][o] for i in range(len(XC)))
                - sum(μA[r]*YA[r][o] for r in range(len(YA)))
                - sum(μC[r]*YC[r][o] for r in range(len(YC))) == 1)
    for j in range(n):
        M.addConstr(sum(vA[i]*XA[i][j] for i in range(len(XA)))
                    + sum(vC[i]*XC[i][j] for i in range(len(XC)))
                    - sum(μA[r]*YA[r][j] for r in range(len(YA)))
                    - sum(μB[r]*YB[r][j] for r in range(len(YB)))
                    - sum(μC[r]*YC[r][j] for r in range(len(YC))) >= 0)
    M.addConstr(sum(μB[r]*YB[r][o] for r in range(len(YB))) == 1)    
    M.Params.DualReductions = 0
    M.optimize()
    status = M.status
    if status == GRB.Status.OPTIMAL:
        epsilonP[o] = M.objVal
    elif status == GRB.Status.INFEASIBLE:
        epsilonP[o] = "Infeasible"
    elif status == GRB.Status.UNBOUNDED:
        epsilonP[o] = "Unbounded"

epsilonN = {}
for o in range(n):
    M = Model()
    vA,vC,μA,μB,μC = {},{},{},{},{}
    for i in range(len(XA)):
        vA[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for i in range(len(XC)):
        vC[i] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YA)):
        μA[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YB)):
        μB[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    for r in range(len(YC)):
        μC[r] = M.addVar(lb = 0, ub = GRB.INFINITY, vtype = "c")
    M.modelSense = GRB.MAXIMIZE
    M.update()
    M.setObjective(sum(vA[i]*XA[i][o] for i in range(len(XA))) - sum(μA[r]*YA[r][o] for r in range(len(YA))))
    M.addConstr(sum(vA[i]*XA[i][o] for i in range(len(XA)))
                + sum(vC[i]*XC[i][o] for i in range(len(XC)))
                - sum(μA[r]*YA[r][o] for r in range(len(YA)))
                - sum(μC[r]*YC[r][o] for r in range(len(YC))) == 1)
    for j in range(n):
        M.addConstr(sum(vA[i]*XA[i][j] for i in range(len(XA)))
                    + sum(vC[i]*XC[i][j] for i in range(len(XC)))
                    - sum(μA[r]*YA[r][j] for r in range(len(YA)))
                    - sum(μB[r]*YB[r][j] for r in range(len(YB)))
                    - sum(μC[r]*YC[r][j] for r in range(len(YC))) >= 0)
    M.addConstr(sum(μB[r]*YB[r][o] for r in range(len(YB))) == 1)    
    M.Params.DualReductions = 0
    M.optimize()
    status = M.status
    if status == GRB.Status.OPTIMAL:
        epsilonN[o] = M.objVal
    elif status == GRB.Status.INFEASIBLE:
        epsilonN[o] = "Infeasible"
    elif status == GRB.Status.UNBOUNDED:
        epsilonN[o] = "Unbounded"

wb = xlwt.Workbook()
ws = wb.add_sheet("Sheet1")
ws.write(0,0,"epsilonP")
ws.write(0,1,"epsilonN")
ws.write(0,2,"RHE")
ws.write(0,3,"LHE")
for j in range(n):
    ws.write(j+1,0,epsilonP[j])
    ws.write(j+1,1,epsilonN[j])
    if epsilonP[j] == "Infeasible" or epsilonP[j] == "Unbounded":
        ws.write(j+1,2,"Undefined")
    else:
        ws.write(j+1,2,epsilonP[j])
    if epsilonN[j] == "Infeasible" or epsilonN[j] == "Unbounded":
        ws.write(j+1,3,"Undefined")
    else:
        ws.write(j+1,3,epsilonN[j])
ask6 = input("Please enter the name under which you want to save your elasticity output file:")
save_as = ask6 + ".xls"
wb.save(save_as)

