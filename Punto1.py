from openpyxl import*
from gurobipy import*

# Carga de Parametros
book = load_workbook("Tarea 2-202420.xlsx")

Sheet=book["Opti-drones"]

#Definicion de Parametros
b = {} #Beneficios
for fila in range(2,27):
    lugar = Sheet.cell(fila,1).value
    print(lugar)
    #for columna in range (2,8):
        
      #  calle = Sheet.cell(1,columna).value
       # b[(lugar,calle)]=Sheet.cell(fila,columna).value

print(b)
        
"""
#Conjuntos
S = list() #Sedes
for fila in range(2,6):
    sede = beneficiosSheet.cell(fila,1).value
    S.append(sede)
#print(S)

F = list() #Facultades
for columna in range(2,8):
    facultad = beneficiosSheet.cell(1,columna).value
    F.append(facultad)
#print(F)



#Modelo de Optimización
m = Model("Punto1")

#Definicion de variables de decisión
x = m.addVars(F,S,vtype=GRB.BINARY, name = "x")

y = m.addVars(F,S,F,S,vtype=GRB.CONTINUOUS, name = "y")

#Restricciones del problema

#1. Todas las facutades se encuentran en una sede
for f in F:
    m.addConstr(quicksum(x[f,s] for s in S) == 1)

#2. Todas las sedes tienen al menos una facultad
for s in S:
    m.addConstr(quicksum(x[f,s] for f in F) >= 1)
    
#3. Todas las sedes tienen maximo tres facultades
for s in S:
    m.addConstr(quicksum(x[f,s] for f in F) <= 3)
    
#4.	La variable auxiliar yfsae siempre será menor o igual que xfs
for s in S:
    for f in F:
        for e in S:
            for a in F:
                m.addConstr(y[f,s,a,e] <= x[f,s])


#5.	La variable auxiliar yfsae siempre será menor o igual que xae
for s in S:
    for f in F:
        for e in S:
            for a in F:
                m.addConstr(y[f,s,a,e] <= x[a,e])
                

#6.	Restricción que cubre el ultimo caso de la tabla de verdad
for s in S:
    for f in F:
        for e in S:
            for a in F:
                m.addConstr(y[f,s,a,e] >= x[f,s]+x[a,e]-1)
        
    

#Funcion Objetivo
m.setObjective(quicksum(x[f,s]*b[f,s] for f in F for s in S)
               -quicksum(y[f,s,a,e]*c[f,a]*k[s,e] for f in F for s in S for a in F for e in S),GRB.MAXIMIZE)

#print(quicksum(x[f,s]*b[f,s] for f in F for s in S))
#
m.update()
m.setParam("Outputflag",0)
m.optimize()
z = m.getObjective().getValue()

#Imrpimir resultados en consola
print(z)

for f,s, in x.keys():
    if x[f,s].x>0:
        print("Facultad ",f," en la sede ",s)
    
    
    
"""