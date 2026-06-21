from core.hardware import inventario_hardware

from core.database import (
    guardar_equipo,
    guardar_colaborador,
    buscar_colaborador
)



def registrar_inventario(datos_colaborador):


    colaborador = buscar_colaborador(
        datos_colaborador["tipo_documento"],
        datos_colaborador["documento"]
    )


    if colaborador:

        colaborador_id = colaborador[0]

        estado = "existente"


    else:

        colaborador_id = guardar_colaborador(
            datos_colaborador
        )

        estado = "nuevo"



    equipo = inventario_hardware()


    equipo["colaborador_id"] = colaborador_id


    guardar_equipo(
        equipo
    )


    return {
        "estado": estado,
        "equipo": equipo
    }