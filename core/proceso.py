import threading


def ejecutar_async(funcion, callback):

    hilo = threading.Thread(
        target=funcion
    )

    hilo.daemon = True

    hilo.start()