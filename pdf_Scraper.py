import tabula as tb
import pandas as pd
import pathlib
from datetime import datetime, timedelta
import numpy as np

#pyuic5 -x mainWindowsUI.ui -o mainWindowsUI.py


#SI(G5>=8,5/24;G5-8/24;0)
def crearReporte(files_raw, pathDestino)->list:
    Files = files_raw
    
    # for j in range(0, len(files_raw)):
    #     if(pathlib.Path(files_raw[j]).suffix == '.pdf'):
    #         Files.append(files_raw[j])
    
    print(Files)
    cedulas = []
    nombres = []
    estados = []
    horarioPreEstablecidoEntrada = []
    horarioPreEstablecidoSalida = []
    primerMarca = []
    ultimaMarca = []
    cantHorasPorFuncionario = []
    totalHorasExtras = 0
    acumuladorDeHorasPorFunc = [datetime.strptime("00:00", "%H:%M")]*30
    acumuladorDeHorasExtrasPorFunc = [datetime.strptime("00:00", "%H:%M")]*30
    serviciosConHorarioFijo = ["6390", "6391", "6393", "6395"]
    cedulasConHorarioFelxEnOfiHoraFijo = ["1693648", "3954879"]

    mapFiles = {}
    for i in range(0, len(Files)):  
   
        numeroDeServicioRaw = tb.read_pdf(f''+Files[i], pages = '1', area = (130, 0, 170, 650), columns = [650], pandas_options = {'header': None}, stream = True)[0]
        numServ = int(str(numeroDeServicioRaw[0]).split(":")[1].split("-")[0].strip())
        if numServ in mapFiles:
            mapFiles[numServ].append(Files[i])
        else:
            mapFiles[numServ] = [Files[i]]

    dfOutput = []
    mapFuncionarios = {}
    mapDeltas = {}
    
    
    
    
    def procesarComoHorarioFijo(acumuladorDeHorasPorFunc, ultimaMarca, primerMarca, horarioPreEstablecidoEntrada, horarioPreEstablecidoSalida, i):
        #Procesar como oficina sin horario fijo
        horaEntrada = datetime.strptime(horarioPreEstablecidoEntrada[i], "%H:%M")
        diferenciaTiempoDePrimerMarca = str((horaEntrada - primerMarca[i]))
        print(diferenciaTiempoDePrimerMarca)
        #Uso que la resta de hora tiene la palabra day si primerMarca es mayor que horaEntrada. Y uso esto como flag
        if not diferenciaTiempoDePrimerMarca.__contains__("day"):
            horas, minutos, segundos = map(str, diferenciaTiempoDePrimerMarca.split(":"))
            if(int(minutos) >= 30 or int(horas) > 0):
                delta = timedelta(hours=horas, minutes=minutos)
                acumuladorDeHorasPorFunc += delta
                
        horaSalida = datetime.strptime(horarioPreEstablecidoSalida[i], "%H:%M")
        diferenciaTiempoDeUltimaMarca = str((ultimaMarca[i] - horaSalida))
        print(diferenciaTiempoDeUltimaMarca)
        print("-------------")
        #Uso que la resta de hora tiene la palabra day si ultimaMarca es menor que horaSalida. Y uso esto como flag
        if not diferenciaTiempoDeUltimaMarca.__contains__("day"):
            horas, minutos, segundos = map(str, diferenciaTiempoDeUltimaMarca.split(":"))
            if(int(minutos) >= 30 or int(horas) > 0):
                delta = timedelta(hours=horas, minutes=minutos)
                acumuladorDeHorasPorFunc += delta

    def procesarComoHorarioFlexible(acumuladorDeHorasExtrasPorFunc, cantHorasPorFuncionario, cedulas, tiempoTrabajado, horasTrabajadas, minutosTrabajados, mapDeltas, i):
        if horasTrabajadas > cantHorasPorFuncionario[i] or (horasTrabajadas == cantHorasPorFuncionario[i] and int(minutosTrabajados) >= 30):
            horas, minutos, seconds = map(str, str(tiempoTrabajado).split(":"))
            aux = datetime.strptime(str(cantHorasPorFuncionario[i])+":00", "%H:%M")
            aux2 = datetime.strptime(horas+":"+minutos, "%H:%M")
            tiempoExtra = str(abs(aux - aux2)).split(" ")[0]
            delta = timedelta(hours=int(tiempoExtra.split(":")[0]), minutes=int(tiempoExtra.split(":")[1]))
            acumuladorDeHorasExtrasPorFunc[i] += delta
            if(int(cedulas[i]) in mapDeltas):
                mapDeltas[cedulas[i]].append(str(delta))
            else:
                mapDeltas[cedulas[i]] = [str(delta)]
        else:
            if(int(cedulas[i]) in mapDeltas):
                mapDeltas[cedulas[i]].append("00:00")
            else:
                mapDeltas[cedulas[i]] = ["00:00"]

    
    
    for j in mapFiles:
        files = mapFiles[j]

        dfOut = pd.DataFrame(columns=['Cedula', 'Nombre', 'Horas trabajadas en la semana', 'cantidad de horas minimas en la semana', 'tiempo extra', "horas extras", "horas trabjadas", "deltas"])
        for i in range(0, len(files)):
            
            numeroDeServicioRaw = tb.read_pdf(f''+files[i], pages = '1', area = (130, 0, 170, 650), columns = [650], pandas_options = {'header': None}, stream = True)[0]
                                                                                                    #0    1    2    3    4    5    6   7    8
            data_frame = tb.read_pdf(f''+files[i], pages = '1', area = (200, 0, 740, 650), columns = [50, 230, 280, 310, 345, 430, 440, 490, 650], pandas_options = {'header': None}, stream = True)[0]
            
            cedulasRaw = data_frame[0]
            nombresRaw = data_frame[1]
            horarioPreEstablecidoEntradaRaw = data_frame[3]
            horarioPreEstablecidoSalidaRaw = data_frame[4]
            estados = data_frame[5]
            primerMarcaRaw = data_frame[7] if 7 in data_frame.columns else [np.nan]*len(nombresRaw)
            ultimaMarcaRaw = data_frame[8] if 8 in data_frame.columns else [np.nan]*len(nombresRaw)


            # print(nombresRaw)
            # print(cedulasRaw)
            # print(horarioPreEstablecidoEntradaRaw)
            # print(horarioPreEstablecidoSalidaRaw)
            # print(primerMarcaRaw)
            # print(ultimaMarcaRaw)


            cedulas = []
            nombres = []
            horarioPreEstablecidoEntrada = []
            horarioPreEstablecidoSalida = []
            primerMarca = []
            ultimaMarca = []
            for i in range(0, len(nombresRaw)):
                if(not pd.isna(nombresRaw[i])):
                    nombres.append(nombresRaw[i])
                    cedulas.append(cedulasRaw[i])
                    horarioPreEstablecidoEntrada.append(horarioPreEstablecidoEntradaRaw[i])
                    horarioPreEstablecidoSalida.append(horarioPreEstablecidoSalidaRaw[i])
                    primerMarca.append(primerMarcaRaw[i])
                    ultimaMarca.append(ultimaMarcaRaw[i])   
            
            # print(nombres)
            # print(cedulas)
            # print(horarioPreEstablecidoEntrada)
            # print(horarioPreEstablecidoSalida)
            # print(primerMarca)
            # print(ultimaMarca)
            cantHorasPorFuncionario = []
            for i in range(0, len(nombres)):
                if(not pd.isna(nombres[i])):
                    marcaSalida = datetime.strptime(horarioPreEstablecidoSalida[i], "%H:%M")
                    marcaEntrada = datetime.strptime(horarioPreEstablecidoEntrada[i], "%H:%M")
                    cantHorasDeTrabajo = marcaSalida - marcaEntrada
                    cantHorasPorFuncionario.append(int(str(cantHorasDeTrabajo).split(":")[0]))


            for i in range(0, len(nombres)):
                if(not pd.isna(nombres[i])):
                    horasTrabajadas = 0
                    if(not pd.isna(primerMarca[i])): 
                        primerMarca[i] = datetime.strptime(primerMarca[i], "%H:%M")
                    if(not pd.isna(ultimaMarca[i])):
                        ultimaMarca[i] = datetime.strptime(ultimaMarca[i], "%H:%M")
                    tiempoTrabajado = ultimaMarca[i]-primerMarca[i] if(not pd.isna(primerMarca[i]) and not pd.isna(ultimaMarca[i])) else 0

                    if(int(cedulas[i]) in mapFuncionarios):
                        mapFuncionarios[cedulas[i]].append(str(tiempoTrabajado))
                    else:
                        mapFuncionarios[cedulas[i]] = [str(tiempoTrabajado)]
                    if(not pd.isna(tiempoTrabajado) and tiempoTrabajado != 0):
                        horasTrabajadas = str(tiempoTrabajado).split(":")[0]
                        minutosTrabajados = str(tiempoTrabajado).split(":")[1]
                        horasTrabajadas = int(horasTrabajadas)
                        
                        

                        acumuladorDeHorasPorFunc[i] += tiempoTrabajado
                        if not serviciosConHorarioFijo.__contains__(str(j)):
                            procesarComoHorarioFlexible(acumuladorDeHorasExtrasPorFunc, cantHorasPorFuncionario, cedulas, tiempoTrabajado, horasTrabajadas, minutosTrabajados, mapDeltas, i)
                        else:
                            if(not cedulasConHorarioFelxEnOfiHoraFijo.__contains__(cedulas[i])):
                                procesarComoHorarioFijo(acumuladorDeHorasPorFunc, ultimaMarca, primerMarca, horarioPreEstablecidoEntrada, horarioPreEstablecidoSalida, i)
                            else:
                                procesarComoHorarioFlexible(acumuladorDeHorasExtrasPorFunc, cantHorasPorFuncionario, cedulas, tiempoTrabajado, horasTrabajadas, minutosTrabajados, mapDeltas, i)

        for i in range(0, len(nombres)):
            if(not pd.isna(nombres[i])):
                cantFiles = len(files)
                #tiempoExtra = (acumuladorDeHorasPorFunc[i] - datetime.strptime(str(int(cantHorasPorFuncionario[i])*cantFiles)+":00", "%H:%M")) 
                aux = str(acumuladorDeHorasPorFunc[i]).split(" ")[1]
                cantDias = acumuladorDeHorasPorFunc[i].day - 1
                aux2 = aux.split(":")
                diasAHoras = int(cantDias)*24
                k = int(aux2[0])+diasAHoras
                tiempo1 = str(k)+":"+aux2[1]
                tiempo2 = str(int(cantHorasPorFuncionario[i])*cantFiles)+":00"   

                
                soloTiempo = str(acumuladorDeHorasExtrasPorFunc[i]).split(" ")[1]
                horasExtras, minutosExtras, secondsExtras = map(int, str(soloTiempo).split(":"))
                sobremeditacion = (horasExtras*60 + minutosExtras)/60
                x = int(sobremeditacion)  # Parte entera
                y = sobremeditacion - x  # Parte decimal

                redondeado = x + 1 if y >= 0.5 else x
                #print(resultado_formateado)
                myDict = {'Cedula':cedulas[i], 
                          'Nombre':nombres[i], 
                          'Horas trabajadas en la semana':str(tiempo1), 
                          'cantidad de horas minimas en la semana':str(int(cantHorasPorFuncionario[i])*cantFiles)+":00",
                          'tiempo extra':str(acumuladorDeHorasExtrasPorFunc[i]).split(" ")[1],
                          'horas extras':str(redondeado),
                          'horas trabjadas':str(mapFuncionarios[cedulas[i]]),
                          'deltas':str(mapDeltas[cedulas[i]])}
                nueva_fila = pd.DataFrame([myDict])
                dfOut = pd.concat([dfOut, nueva_fila], ignore_index=True)

        #nueva_fila = pd.DataFrame("Total horas extras: " + str(totalHorasExtras))
        #dfOut = pd.concat([dfOut, nueva_fila], ignore_index=True)

        dfOutput.append(dfOut)

    # for i in range (0, len(dfOutput)):
    #     dfOutput[i].to_csv("reporteHorasExtras"+str(i)+".csv")
    # wb= Workbook()
    # ws=wb.active
    # with pd.ExcelWriter('reporteHorasExtras.xlsx', engine="openpyxl") as writer:
    #     writer.book=wb
    #     writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
    #     for i in dfOutput:
    #         dfOutput[i].to_excel(writer, sheet_name='Hoja'+str(i), index=False)
    #         writer.save()
    
    return dfOutput
    #dfOut.to_excel("reporteHorasExtras.xlsx")