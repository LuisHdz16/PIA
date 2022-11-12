import datetime
import sys
import sqlite3
from sqlite3 import Error
import openpyxl
import os

def select_clientes():
    try:
        with sqlite3.connect("pia.db") as conn_clientes_select:
            mi_cursor = conn_clientes_select.cursor()
            mi_cursor.execute("SELECT id_cliente, nombre_cliente FROM clientes")
            clientes_registros = mi_cursor.fetchall()
        conn_clientes_select.close()
        return clientes_registros      
    except Error as e:
        print(e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_clientes_por_id(cliente_id):
    try:
        with sqlite3.connect("pia.db") as conn_clientes_select_id:
            mi_cursor = conn_clientes_select_id.cursor()
            mi_cursor.execute(f"SELECT  nombre_cliente FROM clientes WHERE id_cliente={cliente_id}")
            clientes_registros_id = mi_cursor.fetchall()
        conn_clientes_select_id.close()
        return clientes_registros_id      
    except Error as e:
        print(e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_salas():
    try:
        with sqlite3.connect("pia.db")as conn_salas_select:
            mi_cursor = conn_salas_select.cursor()
            mi_cursor.execute("SELECT id_sala, nombre_sala, cupo FROM salas")
            registros_salas = mi_cursor.fetchall()
        conn_salas_select.close()
        return registros_salas      
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_salas_por_id(salas_id):
    try:
        with sqlite3.connect("pia.db")as conn_salas_select_id:
            mi_cursor = conn_salas_select_id.cursor()
            mi_cursor.execute(f"SELECT nombre_sala, cupo FROM salas WHERE id_sala={salas_id}")
            registros_salas = mi_cursor.fetchall()
        conn_salas_select_id.close()
        return registros_salas      
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_turnos():
    try:
        with sqlite3.connect("pia.db")as conn_turnos_select:
            mi_cursor = conn_turnos_select.cursor()
            mi_cursor.execute("SELECT id_turno, nombre_turno FROM turnos")
            registros_turnos = mi_cursor.fetchall()
        conn_turnos_select.close()
        return registros_turnos      
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_turnos_por_id(turno_id):
    try:
        with sqlite3.connect("pia.db")as conn_turnos_select_id:
            mi_cursor = conn_turnos_select_id.cursor()
            mi_cursor.execute(f"SELECT nombre_turno FROM turnos WHERE id_turno={turno_id}")
            registros_turnos = mi_cursor.fetchall()
        conn_turnos_select_id.close()
        return registros_turnos      
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_reservas():
    try:
        with sqlite3.connect("pia.db")as conn_reservas_select:
            mi_cursor = conn_reservas_select.cursor()
            mi_cursor.execute("SELECT folio, fecha, id_cliente, id_sala, nombre_evento ,id_turno FROM reservas")
            registros_reservas = mi_cursor.fetchall()
        conn_reservas_select.close()
        return registros_reservas      
    except Error as e:
        print (e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

def select_reservas_por_id(reserva_id):
    try:
        with sqlite3.connect("pia.db") as conn_reservas_select_id:
            mi_cursor = conn_reservas_select_id.cursor()
            mi_cursor.execute(f"SELECT folio, fecha, id_cliente, id_sala, nombre_evento ,id_turno FROM reservas WHERE folio={reserva_id}")
            registros_reservas = mi_cursor.fetchall()
        conn_reservas_select_id.close()
        return registros_reservas      
    except Error as e:
        print(e)
    except Exception:
        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
    

def puede_ser_tipo_fecha(fecha):
    try:
        return datetime.datetime.strptime(fecha, "%d/%m/%Y").date()
    except Exception:
        return False

def puede_ser_int(valor):
    try:
        return int(valor)
    except Exception:
        return False

if os.path.isfile("pia.db"):
    print("\nSe detecto que hay datos previos.")
else:
    try:
        with sqlite3.connect("pia.db") as conn_creacion_tablas:
            mi_cursor = conn_creacion_tablas.cursor()

            mi_cursor.execute('''CREATE TABLE IF NOT EXISTS clientes (
                                    id_cliente INTEGER PRIMARY KEY AUTOINCREMENT, 
                                    nombre_cliente TEXT NOT NULL);''')

            mi_cursor.execute('''CREATE TABLE IF NOT EXISTS salas (
                                    id_sala INTEGER PRIMARY KEY AUTOINCREMENT, 
                                    nombre_sala TEXT NOT NULL, 
                                    cupo INTEGER NOT NULL);''')

            mi_cursor.execute('''CREATE TABLE IF NOT EXISTS turnos (
                                    id_turno INTEGER PRIMARY KEY AUTOINCREMENT, 
                                    nombre_turno TEXT NOT NULL);''')

            mi_cursor.execute('''CREATE TABLE IF NOT EXISTS reservas (
                                    folio INTEGER PRIMARY KEY AUTOINCREMENT, 
                                    fecha timestamp NOT NULL,
                                    id_sala INTEGER NOT NULL, 
                                    id_cliente INTEGER NOT NULL, 
                                    nombre_evento TEXT NOT NULL, 
                                    id_turno INTEGER NOT NULL, 
                                    FOREIGN KEY(id_cliente) REFERENCES clientes(id_cliente), 
                                    FOREIGN KEY(id_sala) REFERENCES salas(id_sala), 
                                    FOREIGN KEY(id_turno) REFERENCES turnos(id_turno));''')

            mi_cursor.execute("INSERT INTO turnos (nombre_turno) VALUES('Matutino')")
            mi_cursor.execute("INSERT INTO turnos (nombre_turno) VALUES('Vespertino')")
            mi_cursor.execute("INSERT INTO turnos (nombre_turno) VALUES('Nocturno')")
            
            print("\nEs la primera vez que se ejecuta el programa.")
        conn_creacion_tablas.close()
    except Error as e:
        print(e)

while True:

    print(f"\n{'-' * 60}")
    print(f"{'MENÚ PRINCIPAL':^60}")
    print("-" * 60)
    print(f"[A] RESERVACIONES.\n")
    print(f"[B] REPORTES.\n")
    print(f"[C] REGISTRAR NUEVO CLIENTE.\n")
    print(f"[D] REGISTRAR NUEVA SALA.\n")
    print(f"[E] Salir")
    print(f"{'-' * 60}\n")

    opcion_menu = input("Seleccione una opción: ")

    if opcion_menu.upper() == "A":

        while True:

            print(f"\n{'-' * 60}")
            print(f"{'MENÚ RESERVAS':^60}")
            print("-" * 60)
            print(f"[A] REGISTRAR NUEVA RESERVACIÓN\n")
            print(f"[B] MODIFICAR DESCRIPCION DE UNA RESERVACIÓN\n")
            print(f"[C] CONSULTAR DISPONIBIBLIDAD DE SALAS PARA UNA FECHA\n")
            print(f"[D] ELIMINAR UNA RESERVACIÓN\n")
            print(f"[E] VOLVER AL MENÚ PRINCIPAL")
            print(f"{'-' * 60}\n")

            opcion_menu_reservas = input("Seleccione una opción: ")

            if opcion_menu_reservas.upper() == "A":

                while True:

                    clientes_registros = select_clientes()

                    print(f"\n{'Clientes registrados':^40}")
                    print("-" * 40)
                    print("Clave\t\tCliente")
                    print("-" * 40)
                    for clave, nombre in clientes_registros:
                        print(f"{clave}\t\t{nombre}")
                    print("-" * 40)

                    reserva_clave_del_cliente_capturada = input("\nIngrese la clave del cliente: ")

                    if reserva_clave_del_cliente_capturada.strip() == "":
                        print("\nLa clave del cliente no puede omitirse. Intente de nuevo.")
                        continue

                    reserva_clave_del_cliente = puede_ser_int(reserva_clave_del_cliente_capturada)
                    
                    if reserva_clave_del_cliente == False:
                        print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                        continue

                    clientes_registros_por_id = select_clientes_por_id(reserva_clave_del_cliente)

                    if clientes_registros_por_id:
                        print(f"\nCliente {clientes_registros_por_id[0][0]} puede continuar con la resevación.")
                        break
                    else:
                        print("\nNo se encontró la clave del cliente. Intente de nuevo.")
                        continue

                while True:
                    
                    registros_salas = select_salas()
                    
                    print(f"\n{'Salas registradas':^40}")
                    print("-" * 40)
                    print(f"{'Clave':<10}{'Nombre':<25}{'Cupo'}")
                    print("-" * 40)
                    for clave, nombre, cupo in registros_salas:
                        print(f"{clave:<10}{nombre:<25}{cupo}")
                    print("-" * 40)

                    clave_de_sala_reservas_capturada = input("\nIngrese el número de sala: ")

                    if clave_de_sala_reservas_capturada.strip() == "":
                        print("\nEl número de sala no puede omitirse. Intente de nuevo.")
                        continue

                    clave_de_sala_reservas = puede_ser_int(clave_de_sala_reservas_capturada)

                    if clave_de_sala_reservas == False:
                        print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                        continue

                    registros_salas_por_id = select_salas_por_id(clave_de_sala_reservas)

                    if registros_salas_por_id:
                        print(f"\nSala {registros_salas_por_id[0][0]} seleccionada.")
                        break
                    else:
                        print("\nNo se encontró el numero de la sala. Intente de nuevo.")
                        continue

                while True:

                    fecha_actual = datetime.date.today()

                    fecha_reservacion_capturada = input("\nIngrese la fecha de reservación que desea con el formato (dd/mm/aaaa): ")

                    if fecha_reservacion_capturada.strip() == "":
                        print("\nLa fecha no puede omitirse. Intente de nuevo.")
                        continue

                    fecha_reservacion = puede_ser_tipo_fecha(fecha_reservacion_capturada)

                    if fecha_reservacion == False:
                        print("\nEl dato proporcionado no es de tipo de fecha correcta. Intente de nuevo.")
                        continue

                    resta_fecha = fecha_reservacion - fecha_actual

                    if resta_fecha.days < 2 and resta_fecha.days >= 0:
                        print("\nLa reservación debe hacerse dos días antes del día elegido.")
                        continue
                    elif resta_fecha.days < 0:
                        print("\nEsa fecha ya pasó.")
                        continue
                    
                    break

                while True:

                    registros_turnos = select_turnos()
                    
                    print(f"\n{'Turnos posibles':^40}")
                    print("-" * 40)
                    print("Clave\t\tTurno")
                    print("-" * 40)
                    for clave, turno in registros_turnos:
                        print(f"{clave}\t\t{turno}")
                    print("-" * 40)

                    turno_reservacion_capturada = input("\nIngrese la clave del turno: ")

                    if turno_reservacion_capturada.strip() == "":
                        print("\nLa clave del turno no puede omitirse. Intente de nuevo.")
                        continue
                    
                    turno_reservacion = puede_ser_int(turno_reservacion_capturada)

                    if turno_reservacion == False:
                        print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                        continue
                    
                    registros_turnos_por_id = select_turnos_por_id(turno_reservacion)

                    if registros_turnos_por_id:
                        print(f"\nTurno {registros_turnos_por_id[0][0]} seleccionada.")
                        break
                    else:
                        print("\nNo se encontró la clave del turno. Intente de nuevo.")
                        continue
                
                while True:
                                    
                    registros_reservas = select_reservas()
                    
                    for folio, fecha, id_cliente, id_sala, nombre_evento ,id_turno in registros_reservas:
                        if fecha_reservacion.strftime("%d/%m/%Y") == fecha and clave_de_sala_reservas == id_sala and turno_reservacion == id_turno:
                            print(f"\nSala y turno ocupados en la fecha proporcionada. Intente de nuevo.")
                            break
                        else:
                            continue
                    else:

                        while True:

                            nombre_evento = input("\nIngrese el nombre del evento: ")

                            if nombre_evento.strip() == "":
                                print("\nEl nombre del evento no puede omitirse. Intente de nuevo.")
                                continue

                            break
                        try:
                            with sqlite3.connect("pia.db")as conn_reservas:
                                mi_cursor = conn_reservas.cursor()
                                insert_reservas = (fecha_reservacion.strftime("%d/%m/%Y"), reserva_clave_del_cliente, clave_de_sala_reservas, nombre_evento, turno_reservacion)
                                mi_cursor.execute("INSERT INTO reservas (fecha, id_cliente, id_sala, nombre_evento, id_turno) VALUES (?,?,?,?,?)", insert_reservas)
                                print("\nReservación completada correctamente.")
                            conn_reservas.close()      
                        except Error as e:
                            print (e)
                        except Exception:
                            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                    break

            elif opcion_menu_reservas.upper() == "B":
                while True:

                    try:
                        with sqlite3.connect("pia.db")as conn_modificar_nombre:
                            mi_cursor = conn_modificar_nombre.cursor()
                            mi_cursor.execute("SELECT folio, nombre_evento FROM reservas")
                            registros_modificar_nombre = mi_cursor.fetchall()
                            print(f"\n{'Modificar nombre':^50}")
                            print("-" * 50)
                            print(f"{'Folio':<10}{'Nombre del evento'}")
                            print("-" * 50)
                            for folio, nombre_evento in registros_modificar_nombre:
                                print(f"{folio:<10}{nombre_evento}")
                            print("-" * 50)
                        conn_modificar_nombre.close()      
                    except Error as e:
                        print (e)
                    except Exception:
                        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                    
                    folio_editar_nombre = input("\nIngrese el folio de la reservación que quiere editar el nombre: ")

                    if folio_editar_nombre.strip() == "":
                        print("\nEl folio no puede omitirse. Intente de nuevo.")
                        continue

                    folio_editar_nombre_int = puede_ser_int(folio_editar_nombre)

                    if folio_editar_nombre_int == False:
                        print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                        continue
                    
                    for folio, nombre_del_evento in registros_modificar_nombre:
                        if folio == folio_editar_nombre_int:
                            while True:
                                nombre_actualizado = input("\nIngrese el nuevo nombre del evento: ")

                                if nombre_actualizado.strip() == "":
                                    print("\nEl nuevo nombre del evento no puede omitirse. Intente de nuevo.")
                                    continue
                                break
                            try:
                                with sqlite3.connect("pia.db") as conn_modificado_nombre:
                                    mi_cursor = conn_modificado_nombre.cursor()
                                    actualizacion_nombre = (nombre_actualizado,folio_editar_nombre_int)
                                    mi_cursor.execute("UPDATE reservas SET nombre_evento=? WHERE folio=?", actualizacion_nombre)
                                    print("\nNombre del evento editado correctamente.")
                                    break    
                            except Error as e:
                                print (e)
                            except Exception:
                                print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                            
                            break
                    else:
                        print("\nNo existe el folio proporcionado. Intente de nuevo.")
                        continue
                            
                    break
            elif opcion_menu_reservas.upper() == "C":
                while True:
                    
                    listas_ocupadas = list()
                    lista_posibles = list()

                    fecha_para_ver_disponibles_capturada = input("\nIngrese la fecha donde desea ver la disponibilidad dd/mm/aaaa: ")

                    if fecha_para_ver_disponibles_capturada.strip() == "":
                        print("\nLa fecha no puede omitise. Intente de nuevo.")
                        continue

                    fecha_para_ver_disponibles = puede_ser_tipo_fecha(fecha_para_ver_disponibles_capturada)
                    
                    if fecha_para_ver_disponibles == False:
                        print("\nNo es de tipo de fecha correcta. Intente de nuevo.")
                        continue

                    try:
                        with sqlite3.connect("pia.db") as conn_reservas_consultas:
                            mi_cursor = conn_reservas_consultas.cursor()
                            seleccionar_fecha = fecha_para_ver_disponibles.strftime("%d/%m/%Y"),
                            mi_cursor.execute("SELECT id_sala, id_turno FROM reservas WHERE fecha=?", seleccionar_fecha)
                            registros_reservas_consultas = mi_cursor.fetchall()
                        conn_reservas_consultas.close()     
                    except Error as e:
                        print (e)
                    except Exception:
                        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

                    for sala, turno in registros_reservas_consultas:
                            listas_ocupadas.append((sala, turno))
                    
                    reservas_ocupadas = set(listas_ocupadas)

                    
                    registros_turnos = select_turnos()
                    registros_salas = select_salas()
                        

                    for id, sala, cupo in registros_salas:
                        for id_, turno in registros_turnos:
                            lista_posibles.append((id, id_))
                
                    reservas_posibles = set(lista_posibles)

                    reservaciones_disponibles = sorted(list(reservas_posibles - reservas_ocupadas))

                    print(f"\n** Salas disponibles para renta el {fecha_para_ver_disponibles_capturada} **")
                    print(f"\n{'SALA':<20}{'TURNO':>20}")
                    for id_sala, id_turno in reservaciones_disponibles:
                        nombre_sala = select_salas_por_id(id_sala)
                        nombre_turno = select_turnos_por_id(id_turno)
                        print(f"{id_sala},{nombre_sala[0][0]:<20}{nombre_turno[0][0]:>20}")
                    break
            elif opcion_menu_reservas.upper() == "D":
                ciclo_confirmacion = False
                while True:

                    try:
                        with sqlite3.connect("pia.db")as conn_modificar_nombre:
                            mi_cursor = conn_modificar_nombre.cursor()
                            mi_cursor.execute("SELECT folio, nombre_evento FROM reservas")
                            registros_modificar_nombre = mi_cursor.fetchall()
                            print(f"\n{'Resevaciones registradas':^50}")
                            print("-" * 50)
                            print(f"{'Folio':<10}{'Nombre del evento'}")
                            print("-" * 50)
                            for folio, nombre_evento in registros_modificar_nombre:
                                print(f"{folio:<10}{nombre_evento}")
                            print("-" * 50)
                        conn_modificar_nombre.close()      
                    except Error as e:
                        print (e)
                    except Exception:
                        print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")

                    folio_para_eliminar_capturada = input("\nIngrese el folio de la reservación que desea eliminar: ")

                    if folio_para_eliminar_capturada.strip() == "":
                        print("\nEl folio no puede omitirse. Intente de nuevo.")
                        continue

                    folio_para_eliminar = puede_ser_int(folio_para_eliminar_capturada)

                    if folio_para_eliminar == False:
                        print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                        continue
                    
                    registros_reservas = select_reservas_por_id(folio_para_eliminar)

                    if not registros_reservas:
                        print("\nNo existe el folio proporcionado. Intente de nuevo.")
                        continue

                    fecha_de_hoy = datetime.date.today()

                    for folio, fecha, id_cliente, id_sala, nombre_evento ,id_turno in registros_reservas:
                        fecha_procesada = datetime.datetime.strptime(fecha, "%d/%m/%Y").date()
                        resta_fecha = fecha_procesada - fecha_de_hoy
                        if resta_fecha.days < 3:
                            print("\nReservación no eliminada, ya que se debe hacer con 3 dias de anticipación.")
                            break
                        else:
                            print(f"\n{'Datos del folio proporcionado':^60}")
                            print("*" * 60)
                            print(f"Fecha: {fecha_procesada.strftime('%d/%m/%Y')}")
                            print(f"No. de sala: {id_sala}")
                            print(f"Clave del cliente: {id_cliente}")
                            print(f"Evento: {nombre_evento}")
                            print(f"No. de turno: {id_turno}")
                            print("*" * 60)
                            ciclo_confirmacion = True
                            break
                    break

                while ciclo_confirmacion:

                    eliminacion_de_reserva = input("Confirmación de la eliminación(S/N): ")

                    if eliminacion_de_reserva.strip() == "":
                        print("\nNo puede omitirse. Intente de nuevo.")
                        continue

                    if eliminacion_de_reserva.upper() == "S":
                        try:
                            with sqlite3.connect("pia.db") as conn_reservas:
                                mi_cursor = conn_reservas.cursor()
                                folio_eliminado = folio_para_eliminar,
                                mi_cursor.execute("DELETE FROM reservas WHERE folio=?", folio_eliminado)
                                print("\nEliminación de reserva realizada correctamente.")     
                        except Error as e:
                            print (e)
                        except Exception:
                            print(f"Se produjo el siguiente error: {sys.exc_info()[0]}")
                        break
                    elif eliminacion_de_reserva.upper() == "N":
                        print("\nReserva no borrada.")
                        break
                    else:
                        print("\nNo es correcto lo que proporciono. Ingrese S si es si o N si es no.\n")
                        continue

            elif opcion_menu_reservas.upper() == "E":
                break
            else:
                print("\nSeleccione una opción correcta.")
    elif opcion_menu.upper() == "B":
        
        while True:

            print(f"\n{'-' * 60}")
            print(f"{'MENÚ REPORTES':^60}")
            print("-" * 60)
            print(f"[A] REPORTE EN PANTALLA DE RESERVACIONES PARA UNA FECHA\n")
            print(f"[B] EXPORTAR REPORTE TABULAR EN EXCEL\n")
            print(f"[C] VOLVER AL MENÚ PRINCIPAL")
            print(f"{'-' * 60}\n")
            opcion_menu_reportes = input("Seleccione una opción: ")

            if opcion_menu_reportes.upper() == "A":

                while True:

                    consulta_reservaciones = []

                    registros_reservas = select_reservas()

                    fecha_a_consultar_capturada = input("\nIngrese la fecha que desea consultar si hay reservaciones: ")

                    if fecha_a_consultar_capturada.strip() == "":
                        print("\nNo se puede omitir la fecha.")
                        continue
                    
                    fecha_a_consultar = puede_ser_tipo_fecha(fecha_a_consultar_capturada)

                    if fecha_a_consultar == False:
                        print("\nNo es de tipo de fecha correcto. Intente de nuevo.")
                        continue
                    
                    for folio, fecha, id_cliente, id_sala, nombre_evento ,id_turno in registros_reservas:
                        if fecha_a_consultar.strftime("%d/%m/%Y") == fecha:
                            nombre_cliente = select_clientes_por_id(id_cliente)
                            nombre_turno = select_turnos_por_id(id_turno)
                            consulta_reservaciones.append((id_sala, nombre_cliente[0][0], nombre_evento, nombre_turno[0][0]))
                            continue
                            
                    if len(consulta_reservaciones) == 0:
                        print("\nNo hay reservaciones en esa fecha. Intente de nuevo")
                        break

                    print("*" * 100)
                    print(f"{'REPORTE DE RESERVACIONES PARA EL DIA ' + fecha_a_consultar_capturada:^100}")
                    print("*" * 100)
                    print(f"{'SALA':<15}{'CLIENTE':<20}{'EVENTO':<50}TURNO")
                    print("*" * 100)
                    for datos in consulta_reservaciones:
                        print(f"{datos[0]:<15}{datos[1]:<20}{datos[2]:<50}{datos[3]}")
                    print("*" * 100)
                    consulta_reservaciones.clear()
                    break
            elif opcion_menu_reportes.upper() == "B":

                while True:

                    consulta_reservaciones_exportar = []

                    fecha_exportar_capturada = input("\nIngrese la fecha para exportar las reservaciones se tienen: ")

                    if fecha_exportar_capturada.strip() == "":
                        print("\nNo puede omitirse. Intente de nuevo.")
                        continue

                    fecha_exportar = puede_ser_tipo_fecha(fecha_exportar_capturada)

                    if fecha_exportar == False:
                        print("\nNo ese de tipo de fecha correcto. Intente de nuevo.")
                        continue
                    try:
                        with sqlite3.connect("pia.db") as conn_exportacion:
                            mi_cursor = conn_exportacion.cursor()
                            select_exportacion = fecha_exportar.strftime("%d/%m/%Y"),
                            mi_cursor.execute("SELECT id_sala, id_cliente, nombre_evento, id_turno FROM reservas WHERE fecha=?", select_exportacion)
                            registros_para_exportar = mi_cursor.fetchall()
                        conn_exportacion.close()
                    except Error as e:
                        print(e)
                    
                    if not registros_para_exportar:
                        print("\nNo hay resevaciones en la fecha proporcionada. intente de nuevo.")
                        continue
                    
                    for id_sala, id_cliente, nombre_evento, id_turno in registros_para_exportar:
                        nombre_cliente = select_clientes_por_id(id_cliente)
                        nombre_turno = select_turnos_por_id(id_turno)
                        consulta_reservaciones_exportar.append((id_sala, nombre_cliente[0][0], nombre_evento, nombre_turno[0][0]))
                        continue

                    else:
                        libro = openpyxl.Workbook()
                        hoja = libro["Sheet"] 
                        hoja.title = "Primera"
                        hoja.append(("sala", "cliente", "evento", "turno"))
                        for valores in consulta_reservaciones_exportar:
                            hoja.append(valores)
                        libro.save("reservas_tabular.xlsx")
                        print("\nExportado a Excel correctamente.")
                        break
            elif opcion_menu_reportes.upper() == "C":
                break
            else:
                print("\nSeleccione una opción correcta.")
    elif opcion_menu.upper() == "C":

        print()
        print("*" * 10 + "REGISTRO DE CLIENTES" + "*" * 10)
        while True:

            cliente_nombre = input("\nIngrese el nombre del cliente: ")

            if cliente_nombre.strip() == "":
                print("\nEl nombre del cliente no se puede omitir. Intente de nuevo.")
                continue
            try:
                with sqlite3.connect("pia.db") as conn_registro_clientes:
                    mi_cursor = conn_registro_clientes.cursor()
                    cliente = cliente_nombre,
                    mi_cursor.execute("INSERT INTO clientes (nombre_cliente) VALUES(?)", cliente)
                    print("\nSe registro el cliente exitosamente")
                conn_registro_clientes.close()
            except Error as e:
                print(e)
            except:
                print(f"\nSe produjo el siguiente error: {sys.exc_info()[0]}")
            break

    elif opcion_menu.upper() == "D":

        print()
        print("*" * 10 + "REGISTRO DE SALAS" + "*" * 10)
        while True:

            sala_nombre = input("\nIngrese el nombre de sala: ")

            if sala_nombre.strip() == "":
                print("\nNo puede omitirse el nombre de la sala. Intente de nuevo.")
                continue
            break

        while True:

            sala_cupo_capturado = input("\nCupo de la sala: ")

            if sala_cupo_capturado.strip() == "":
                print("\nNo puede omitirse el cupo de la sala. Intente de nuevo.")
                continue

            sala_cupo = puede_ser_int(sala_cupo_capturado)

            if sala_cupo == False:
                print("\nEl dato proporcionado no es de tipo entero. Intente de nuevo.")
                continue

            if sala_cupo <= 0:
                print("\nEl cupo de la sala debe ser mayor a cero. Intente de nuevo.")
                continue

            try:
                with sqlite3.connect("pia.db") as conn_registro_salas:
                    mi_cursor = conn_registro_salas.cursor()
                    sala = sala_nombre, sala_cupo
                    mi_cursor.execute("INSERT INTO salas (nombre_sala, cupo) VALUES(?,?)", sala)
                    print("\nSe registro la sala exitosamente")
                conn_registro_salas.close()
            except Error as e:
                print(e)
            except:
                print(f"\nSe produjo el siguiente error: {sys.exc_info()[0]}")

            break
    elif opcion_menu.upper() == "E":
        break
    else:
        print("\nSeleccione una opción correcta.")