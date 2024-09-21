from openpyxl import*
from gurobipy import*
import networkx as nx
import matplotlib.pyplot as plt
# Carga de Parametros
book = load_workbook("Tarea 2-202420.xlsx")

Sheet=book["Opti-drones"]


#----------------------------------------------
# Definicion de Conjuntos
#----------------------------------------------

# Conjunto de Lugares
L = [] 

for fila in range(2,28):
    lugar = Sheet.cell(fila,1).value
    L.append(lugar)
print(L)

# Conjunto de Arcos posibles
A = [] 
for fila1 in range(2,28):
    lugar1 = Sheet.cell(fila1,1).value
    
    for fila2 in range(2,28):
        lugar2 = Sheet.cell(fila2,1).value
        if lugar1 != lugar2:
            lugar = (lugar1, lugar2)
            A.append(lugar)
print(A)



#----------------------------------------------
# Definicion de Parametros
#----------------------------------------------

# Posicion X (Carrera)
X = {}
for fila in range(2,28):
    lugar =  Sheet.cell(fila,1).value
    posX =  Sheet.cell(fila,2).value
    X[lugar] = posX
print(X)

# Posicion Y (Calle)
Y = {} 
for fila in range(2,28):
    lugar =  Sheet.cell(fila,1).value
    posY =  Sheet.cell(fila,3).value
    Y[lugar] = posY
print(Y)

# Cantidad de fotos requeridas por lugar
F = {}
for fila in range(2,28):
    lugar =  Sheet.cell(fila,1).value
    posY =  Sheet.cell(fila,4).value
    F[lugar] = posY
print(F)

# Distancia Manhattan entre lugares
M = {}
for i in L:
    for j in L:
        if i != j: #Para que no se repitan los lugares
            M[(i,j)] = abs(X[i]-X[j])+abs(Y[i]-Y[j])
print(M)

#----------------------------------------------
# Modelo de Optimización
#----------------------------------------------

#Creación del modelo
m = Model("Tarea2")

#Definicion de variables de decisión
x = m.addVars(A,vtype=GRB.BINARY, name = "x")

#----------------------------------------------
# Restricciones del problema 
#----------------------------------------------


#1
for i in L:
    if i != 'Hangar':
        m.addConstr(quicksum(x[i,j] for j in L if i != j) == 1)

for j in L:
    if j != 'Hangar':
        m.addConstr(quicksum(x[i,j] for i in L if i != j) == 1)
        
#2
m.addConstr(quicksum(x['Hangar',j] for j in L if j != 'Hangar') == 5)
m.addConstr(quicksum(x[i,'Hangar'] for i in L if i != 'Hangar') == 5)
#3
for i in L :
    for j in L :
        if i !=j: 
            m.addConstr(x[i,j] + x[j,i] <= 1)

#F.O
m.setObjective(quicksum(x[a]*M[a] for a in A),GRB.MINIMIZE)
m.update()
#m.setParam("Outputflag",0)
m.optimize()
arcos_reales = []
for i in A:
    if x[i].x > 0:
        #print(i)
        #print(x[i].x)
        arcos_reales.append(i)

G = nx.DiGraph()
G.add_edges_from(arcos_reales)
nx.draw(G, with_labels=True)

#Resultados en tabla con # de Ruta, secuencia, tiempo y fotos
ciclos = list(nx.simple_cycles(G))
#print(ciclos)
#definir cantidad de ciclos
cantidad_ciclos = len(ciclos)

#para cada ciclo definir su secuencia
secuencia = []
for i in range(cantidad_ciclos):
    secuencia.append(ciclos[i])
    print("Ciclo:" + str(i+1) + " " +str(secuencia) + "/n") 
    
#para cada ciclo definir su tiempo
tiempo = []
for i in range(cantidad_ciclos):
    tiempo.append(len(secuencia[i]))
    print("Ciclo:" + str(i+1) + " " +str(tiempo) + "/n")
    
    



#plt.show()

#----------------------------------------------
#Resultados
#----------------------------------------------




