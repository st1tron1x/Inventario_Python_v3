import customtkinter as ctk
from PIL import Image

from gui.estilos import (
    COLOR_FONDO,
    COLOR_CARD
)

from core.inventario import registrar_inventario
from core.database import buscar_colaborador
from core.reportes import exportar_excel
from tkinter import messagebox
from core.config import APP_NAME, APP_VERSION



def generar_reporte():

    try:

        archivo = exportar_excel()

        messagebox.showinfo(
            "Reporte generado",
            f"Archivo creado:\n\n{archivo}"
        )


    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )

def ventana_inventario(parent):

    formulario = ctk.CTkToplevel(parent)

    formulario.title(
        "Nuevo Inventario"
    )

    formulario.geometry(
        "500x650"
    )


    campos = {}


    etiquetas = [

        "Tipo Documento",
        "Documento",
        "Nombre",
        "Cargo",
        "Área",
        "Sede",
        "Correo"

    ]


    for texto in etiquetas:


        ctk.CTkLabel(

            formulario,

            text=texto

        ).pack(
            pady=5
        )



        if texto == "Tipo Documento":


            entrada = ctk.CTkComboBox(

                formulario,

                values=[

                    "CC",
                    "CE",
                    "PASAPORTE",
                    "NIT"

                ]

            )

            entrada.set("CC")


        else:


            entrada = ctk.CTkEntry(

                formulario,

                width=300

            )



        entrada.pack(
            pady=5
        )


        campos[texto] = entrada




    def buscar_documento(event=None):


        documento = campos["Documento"].get()

        tipo = campos["Tipo Documento"].get()


        if len(documento) < 5:

            return



        resultado = buscar_colaborador(

            tipo,

            documento

        )


        if resultado:


            campos["Nombre"].delete(
                0,
                "end"
            )

            campos["Nombre"].insert(
                0,
                resultado[1]
            )



            campos["Cargo"].delete(
                0,
                "end"
            )

            campos["Cargo"].insert(
                0,
                resultado[2]
            )



            campos["Área"].delete(
                0,
                "end"
            )

            campos["Área"].insert(
                0,
                resultado[3]
            )



            campos["Sede"].delete(
                0,
                "end"
            )

            campos["Sede"].insert(
                0,
                resultado[4]
            )



            campos["Correo"].delete(
                0,
                "end"
            )

            campos["Correo"].insert(
                0,
                resultado[5]
            )



    campos["Documento"].bind(
        "<KeyRelease>",
        buscar_documento
    )




    def guardar():


        datos = {


            "tipo_documento":
            campos["Tipo Documento"].get(),


            "documento":
            campos["Documento"].get(),


            "nombre":
            campos["Nombre"].get(),


            "cargo":
            campos["Cargo"].get(),


            "area":
            campos["Área"].get(),


            "sede":
            campos["Sede"].get(),


            "correo":
            campos["Correo"].get()

        }



        try:


            registrar_inventario(
                datos
            )


            messagebox.showinfo(

                "Inventario",

                "Equipo registrado correctamente"

            )


            formulario.destroy()



        except Exception as e:


            messagebox.showerror(

                "Error",

                str(e)

            )




    ctk.CTkButton(

        formulario,

        text="💾 Registrar Equipo",

        width=300,

        height=45,

        command=guardar

    ).pack(

        pady=30

    )



def iniciar_app():


    ventana = ctk.CTk()


    ventana.geometry(
        "1000x650"
    )


    ventana.title(
        f"{APP_NAME} v{APP_VERSION}"
    )


    ventana.resizable(
        False,
        False
    )



    # =============================
    # Fondo corporativo
    # =============================

    try:


        imagen = Image.open(
            "assets/background.png"
        )


        fondo = ctk.CTkImage(

            light_image=imagen,

            dark_image=imagen,

            size=(
                1000,
                650
            )

        )


        fondo_label = ctk.CTkLabel(

            ventana,

            image=fondo,

            text=""

        )


        fondo_label.place(
            x=0,
            y=0,
            relwidth=1,
            relheight=1
        )


    except Exception as e:

        print(
            "No se cargó fondo:",
            e
        )





    # =============================
    # Panel central
    # =============================


    panel = ctk.CTkFrame(

        ventana,

        width=650,

        height=420,

        corner_radius=25,

        fg_color=COLOR_CARD,

        bg_color="transparent"

    )


    panel.place(

        relx=0.5,

        rely=0.5,

        anchor="center"

    )





    titulo = ctk.CTkLabel(

        panel,

        text=
        "INVENTARIO TI\nT&C ASOCIADOS",

        font=(
            "Segoe UI",
            32,
            "bold"
        )

    )


    titulo.pack(
        pady=35
    )





    subtitulo = ctk.CTkLabel(

        panel,

        text=
        "Gestión de activos tecnológicos",

        font=(
            "Segoe UI",
            15
        )

    )


    subtitulo.pack()





    boton_inventario = ctk.CTkButton(

        panel,

        text="➕ Nuevo Inventario",

        width=350,

        height=55,

        corner_radius=15,

        command=lambda:

        ventana_inventario(
            ventana
        )

    )


    boton_inventario.pack(

        pady=25

    )





    boton_reporte = ctk.CTkButton(

        panel,

        text="📊 Generar Reporte Excel",

        width=350,

        height=55,

        corner_radius=15,

        command=generar_reporte

    )


    boton_reporte.pack()







    footer = ctk.CTkLabel(

        ventana,

        text=(
        f"{APP_NAME}\n"
        f"Versión {APP_VERSION}\n"
        "Desarrollado por Ing. Stiven Vanegas Jimenez\n"
        "Derechos reservados de autor"
        ),

        font=(
            "Segoe UI",
            11
        ),

        fg_color="transparent"

    )


    footer.place(

        relx=0.5,

        rely=0.94,

        anchor="center"

    )



    ventana.mainloop()