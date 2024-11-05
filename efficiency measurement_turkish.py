from gurobipy import *
import xlrd
import xlwt
from collections import defaultdict

ask1 = input("Veri setinizde kaç adet KVB vardır?")
n = int(ask1)
ask2 = input("Veri setinizde kaç adet girdi bulunmaktadır?")
m = int(ask2)
ask3 = input("Veri setinizde kaç adet çıktı bulunmaktadır?")
s = int(ask3)
ask4 = input("Hangi veri dosyası ile çalışmak istersiniz?")
file_name = ask4 + ".xls"
book = xlrd.open_workbook(file_name)
ask5 = input("Hangi sayfa ile çalışmak istersiniz?")
sheet = book.sheet_by_index(int(ask5))

x = [[sheet.cell_value(r,c) for r in range(1,sheet.nrows)] for c in range(1, m+1)]
y = [[sheet.cell_value(r,c) for r in range(1,sheet.nrows)] for c in range(m+1, m+s+1)]

virtualx,virtualy = {},{}

# CRS_GO (girdi odaklı CRS çarpan modeli)
crs_go = {}
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
    crs_go[o] = M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["CRS_GO"] = virtual_x
virtualy["CRS_GO"] = virtual_y 

# CRS_ÇO (çıktı odaklı CRS çarpan modeli)
crs_ço = {}
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
    crs_ço[o] = 1/M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["CRS_ÇO"] = virtual_x
virtualy["CRS_ÇO"] = virtual_y 

# VRS_GO (girdi odaklı VRS çarpan modeli)
vrs_go = {}
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
    vrs_go[o] = M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["VRS_GO"] = virtual_x
virtualy["VRS_GO"] = virtual_y 
        
# VRS_ÇO (çıktı odaklı VRS çarpan modeli)
vrs_ço = {}
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
    vrs_ço[o] = 1/M.objVal
    for i in range(m):
        virtual_x[i][o] = v[i].x * x[i][o]
    for r in range(s):
        virtual_y[r][o] = μ[r].x * y[r][o]
virtualx["VRS_ÇO"] = virtual_x
virtualy["VRS_ÇO"] = virtual_y

wb = xlwt.Workbook()
models = {"CRS_GO":crs_go, "CRS_ÇO":crs_ço, "VRS_GO":vrs_go, "VRS_ÇO":vrs_ço}
for model in models:
    ws = wb.add_sheet(model)
    ws.write(0,0,"Skor")
    for i in range(1,m+1):
        string = str(i)
        ws.write(0,i,"Sanal Girdi_"+string)
    for r in range(m+1,m+s+1):
        string = str(r-m)
        ws.write(0,r,"Sanal Çıktı_"+string)
    for j in range(n):
        ws.write(j+1,0,models[model][j])
        for i in range(m):
            ws.write(j+1,i+1,virtualx[model][i][j])
        for r in range(s):
            ws.write(j+1,m+r+1,virtualy[model][r][j])
wb.save("Etkinlik Ölçümü.xls")
