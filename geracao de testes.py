import xlwt, random

indices  = ["Tensão","MQ_1","MQ_2","MQ_3","tempIn","tempOut","Umidade","Outlier_1","Outlier_2","Outlier_3",]
wb = xlwt.Workbook()
ws = wb.add_sheet("Dados")
ws_1 = wb.add_sheet("semInfo")
red_style = xlwt.easyxf('pattern: pattern solid, fore_color red;')
green_style = xlwt.easyxf('pattern: pattern solid, fore_color green;')

for i in range (10): #Escrita dos indices
    ws.write(0, i, indices[i])

for i in range(10):
    for j in range(1, 1000):
        values = [random.uniform(0.0, 5.0), random.randrange(0, 2000, 8), random.randrange(0, 2000, 10), random.randrange(0, 2000, 14), random.randrange(20, 70, 3)/1.3, random.randrange(5, 50, 3)/1.3, random.uniform(0.0, 1.0), random.randrange(0, 2000, 8), random.randrange(0, 2000, 10), random.randrange(0, 2000, 14)]

        if i == 0:
            if values[i] < 2.0 or values[i] > 5.0:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 1:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 2:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 3:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 4:
            if values[i] < 25 or values[i] > 60:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 5:
            if values[i] < 5 or values[i] > 35:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])                                                                

        if i == 6:
            if values[i] < 0.33 or values[i] > 0.85:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])
                
        if i == 7:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 8:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])

        if i == 9:
            if values[i] < 300 or values[i] > 1500:
                ws.write(j, i, values[i], red_style)
                ws_1.write(j, i, values[i])
            else: 
                ws.write(j, i, values[i], green_style)
                ws_1.write(j, i, values[i])                
wb.save("Serie de dados para testagem.xls")


# for i in range (1, 1000):
#     values = [random.uniform(0.0, 5.0), random.randrange(0, 2000, 8), random.randrange(0, 2000, 10), random.randrange(0, 2000, 14), random.randrange(20, 70, 3)/1.3, random.randrange(5, 50, 3)/1.3, random.uniform(0.33, 0.85), random.randrange(0, 2000, 8), random.randrange(0, 2000, 10), random.randrange(0, 2000, 14)]
#     for j in range (10):
#         ws.write(i, j, values[j], red_style)

#Tensao - [0, 5]
#MQ_1 - [0, 2000]
#MQ_2 - [0, 2000]
#MQ_3 - [0, 2000]
#tempIn - [20, 70]
#tempOut - [5, 35]
#umidade - [0.33, 0.85]
#Outlier_1 - [0, 2000]
#Outlier_2 - [0, 2000]
#Outlier_3 - [0, 2000]
