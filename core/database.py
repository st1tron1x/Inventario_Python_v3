#database.py
import sqlite3
import os
import json


#RUTA_DB = "data/inventario.db"
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
RUTA_DB = os.path.join(DATA_DIR, "inventario.db")


def conectar():
    os.makedirs(DATA_DIR, exist_ok=True)
    return sqlite3.connect(RUTA_DB)

def reiniciar_bd():

    conexion = conectar()

    cursor = conexion.cursor()


    cursor.execute(
        "DROP TABLE IF EXISTS equipos"
    )

    cursor.execute(
        "DROP TABLE IF EXISTS colaboradores"
    )


    conexion.commit()

    conexion.close()

def crear_tablas():

    conexion = conectar()

    cursor = conexion.cursor()


    # Tabla equipos

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS equipos (

        id INTEGER PRIMARY KEY AUTOINCREMENT,
        colaborador_id INTEGER,

        hostname TEXT,
        usuario_windows TEXT,

        serial TEXT UNIQUE,

        fabricante TEXT,
        modelo TEXT,

        procesador TEXT,

        memoria TEXT,
        discos TEXT,

        ip TEXT,
        mac TEXT,

        sistema TEXT,

        fecha_registro DATETIME DEFAULT (datetime('now','localtime'))
        

    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS colaboradores (

        id INTEGER PRIMARY KEY AUTOINCREMENT,
                   
        tipo_documento TEXT,
        documento TEXT UNIQUE,

        nombre TEXT NOT NULL,
        cargo TEXT,
        area TEXT,
        sede TEXT,
        correo TEXT,

        fecha_registro DATETIME DEFAULT (datetime('now','localtime'))

    )
    """)


    conexion.commit()

    conexion.close()




def guardar_equipo(datos):


    conexion = conectar()

    cursor = conexion.cursor()



    cursor.execute("""
    
    INSERT INTO equipos
    (
        colaborador_id,
                   
        hostname,
        usuario_windows,
        serial,

        fabricante,
        modelo,

        procesador,

        memoria,
        discos,

        ip,
        mac,

        sistema

    )

    VALUES
    (?,?,?,?,?,?,?,?,?,?,?,?)

    """,

    (

        datos["colaborador_id"],
        
        datos["hostname"],

        datos["usuario"],

        datos["serial"],


        datos["fabricante"],

        datos["modelo"],


        datos["procesador"],


        json.dumps(
            datos["memoria"],
            ensure_ascii=False
        ),

        json.dumps(
            datos["discos"],
            ensure_ascii=False
        ),


        datos["ip"],

        datos["mac"],


        str(datos["sistema"])

    ))



    conexion.commit()

    conexion.close()

def guardar_colaborador(datos):


    conexion = conectar()

    cursor = conexion.cursor()


    cursor.execute(
    """
    INSERT INTO colaboradores
    (
        tipo_documento,
        documento,
        nombre,
        cargo,
        area,
        sede,
        correo
    )

    VALUES (?,?,?,?,?,?,?)

    """,

    (

        datos["tipo_documento"],
        datos["documento"],
        datos["nombre"],
        datos["cargo"],
        datos["area"],
        datos["sede"],
        datos["correo"]

    ))


    conexion.commit()


    colaborador_id = cursor.lastrowid


    conexion.close()


    return colaborador_id

def buscar_colaborador(tipo_documento, documento):

    conexion = conectar()

    cursor = conexion.cursor()


    cursor.execute(
        """
        SELECT id, nombre, cargo, area, sede, correo
        FROM colaboradores
        WHERE tipo_documento = ?
        AND documento = ?
        """,
        (
            tipo_documento,
            documento
        )
    )


    resultado = cursor.fetchone()


    conexion.close()


    return resultado