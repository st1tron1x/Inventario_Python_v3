import socket
import uuid
import subprocess

import psutil
import wmi
import pythoncom



TIPOS_RAM = {

    20: "DDR",
    21: "DDR2",
    24: "DDR3",
    26: "DDR4",
    34: "DDR5"

}



def obtener_wmi():

    return wmi.WMI()



def obtener_hostname():

    return socket.gethostname()





def obtener_usuario():

    try:

        usuarios = psutil.users()

        if usuarios:
            return usuarios[0].name

    except Exception:
        pass


    return "No detectado"





def obtener_serial():


    # WMI clásico

    try:

        for bios in obtener_wmi().Win32_BIOS():

            if bios.SerialNumber:

                return bios.SerialNumber.strip()


    except Exception:

        pass



    # PowerShell CIM Windows 11

    try:


        comando = [

            "powershell",

            "-Command",

            "(Get-CimInstance Win32_BIOS).SerialNumber"

        ]



        resultado = subprocess.check_output(

            comando,

            text=True

        )


        if resultado.strip():

            return resultado.strip()



    except Exception:

        pass



    return "No detectado"







def obtener_equipo():


    datos = {}



    try:


        for pc in obtener_wmi().Win32_ComputerSystem():


            datos["fabricante"] = (
                pc.Manufacturer or ""
            )


            datos["modelo"] = (
                pc.Model or ""
            )


    except Exception:

        pass



    return datos







def obtener_procesador():


    try:


        for cpu in obtener_wmi().Win32_Processor():

            return cpu.Name.strip()



    except Exception:

        pass



    return "No detectado"









def obtener_memoria():


    memorias = []



    try:


        for ram in obtener_wmi().Win32_PhysicalMemory():



            memorias.append(

                {


                "fabricante":

                    ram.Manufacturer
                    or "Unknown",



                "capacidad_gb":

                    round(

                        int(ram.Capacity)

                        / (1024**3),

                        2

                    ),



                "velocidad_mhz":

                    ram.Speed,



                "tipo":

                    TIPOS_RAM.get(

                        ram.SMBIOSMemoryType,

                        "Desconocido"

                    )

                }

            )



    except Exception:

        pass



    return memorias







def detectar_tipo_disco(modelo):


    if not modelo:

        return "Desconocido"



    modelo = modelo.upper()



    if "NVME" in modelo:

        return "NVMe SSD"



    if "SSD" in modelo:

        return "SSD"



    return "HDD"







def obtener_discos():


    discos = []



    try:



        for disk in obtener_wmi().Win32_DiskDrive():



            discos.append(

                {


                "modelo":

                    disk.Model
                    or "Unknown",



                "interfaz":

                    disk.InterfaceType
                    or "Unknown",



                "tipo":

                    detectar_tipo_disco(

                        disk.Model

                    ),



                "capacidad_gb":

                    round(

                        int(disk.Size)

                        /(1024**3),

                        2

                    )

                }

            )



    except Exception:

        pass



    return discos









def obtener_red():


    datos = {}



    try:


        datos["ip"] = socket.gethostbyname(

            socket.gethostname()

        )


    except Exception:


        datos["ip"] = "No detectada"




    mac = uuid.getnode()



    datos["mac"] = ':'.join(

        [

            "{:02x}".format(

                (mac >> i) & 0xff

            )

            for i in range(0,48,8)

        ][::-1]

    )



    return datos







def obtener_windows():


    datos = {



        "nombre":
            "No detectado",


        "version":
            "",


        "build":
            "",


        "arquitectura":
            ""

    }




    try:


        for os in obtener_wmi().Win32_OperatingSystem():


            datos["nombre"] = os.Caption

            datos["version"] = os.Version

            datos["build"] = os.BuildNumber

            datos["arquitectura"] = os.OSArchitecture



    except Exception:

        pass



    return datos







def inventario_hardware():


    pythoncom.CoInitialize()


    try:



        equipo = obtener_equipo()

        red = obtener_red()



        return {



            "hostname":

                obtener_hostname(),



            "usuario":

                obtener_usuario(),



            "serial":

                obtener_serial(),



            "fabricante":

                equipo.get(

                    "fabricante",

                    "No detectado"

                ),



            "modelo":

                equipo.get(

                    "modelo",

                    "No detectado"

                ),



            "procesador":

                obtener_procesador(),



            "memoria":

                obtener_memoria(),



            "discos":

                obtener_discos(),



            "ip":

                red["ip"],



            "mac":

                red["mac"],



            "sistema":

                obtener_windows()

        }



    finally:


        pythoncom.CoUninitialize()