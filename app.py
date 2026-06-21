from core.database import crear_tablas
from gui.ventana import iniciar_app



def main():

    crear_tablas()

    iniciar_app()



if __name__ == "__main__":

    main()