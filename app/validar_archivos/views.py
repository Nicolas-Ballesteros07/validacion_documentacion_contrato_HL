from django.shortcuts import render, redirect
from django.http import JsonResponse
from .forms import ValidacionForm
from .utils import (
    extraer_texto,
    validar_nombre,
    validar_campo_texto,
    validar_monto,
    validar_fecha,
    validar_numero_contacto,
    CAMPOS_POR_TIPO,
    DEFINICIONES_CAMPOS,
    limpiar,
    normalizar_monto,
    formatear_monto,
)
from .buscar_documento_mejorado import extraer_y_validar_documento
from .extractores_por_tipo import extraer_campos_estructurados
from rapidfuzz import fuzz
import re
import json

NUMEROS_PREDETERMINADOS = {
    "900839160",
    "9008391607",
    "39786135",
    "33333333",
}

TIPOS_DOCUMENTOS = {
    "AUXILIO":     "Auxilio",
    "CONDICIONES": "Condiciones",
    "CONTRATO":    "Contrato",
    "MAN_CONF":    "Manual y Confidencialidad",
    "MEDICAS":     "Médicas",
    "RIESGO":      "Riesgo",
}

# Lista de todos los campos que manejamos en el formulario y sesión
CAMPOS_FORMULARIO = [
    "nombre",
    "documento",
    "numero_contacto",
    "cargo",
    "categoria_salarial",
    "salario_basico",
    "beneficio_alimentacion",
    "beneficio_pension",
    "auxilio_localizacion",
    "auxilio_vivienda",
    "condicion_descanso",
    "tipo_contrato",
    "fecha_inicio",
    "proyecto",
    "ciudad",
    "aux_almuerzo",
    "aux_transporte",
    "aux_vivienda",
    "aux_desplazamiento",
]


# =========================================================================
# COMPARADORES SOBRE VALORES YA EXTRAÍDOS (doc → string)
# =========================================================================

def _comparar_nombre(valor_usuario, valor_doc):
    score = fuzz.token_sort_ratio(limpiar(str(valor_usuario)), limpiar(str(valor_doc)))
    if score == 100:
        return {"estado": "CORRECTO", "encontrado": valor_doc, "mensaje": None}
    elif score >= 80:
        return {
            "estado": "SIMILAR",
            "encontrado": valor_doc,
            "mensaje": f'Nombre muy similar (score {score}%). Ingresado: "{valor_usuario}" — Documento: "{valor_doc}"',
        }
    return {
        "estado": "DIFERENTE",
        "encontrado": valor_doc,
        "mensaje": f'Se ingresó "{valor_usuario}" pero el documento tiene "{valor_doc}"',
    }


def _comparar_texto(valor_usuario, valor_doc, umbral=80):
    score = fuzz.token_sort_ratio(limpiar(str(valor_usuario)), limpiar(str(valor_doc)))
    if score == 100:
        return {"estado": "CORRECTO", "encontrado": valor_doc, "mensaje": None}
    elif score >= umbral:
        return {"estado": "SIMILAR", "encontrado": valor_doc, "score": score, "mensaje": None}
    return {
        "estado": "DIFERENTE",
        "encontrado": valor_doc,
        "score": score,
        "mensaje": f'Se ingresó "{valor_usuario}" pero el documento dice "{valor_doc}"',
    }


def _comparar_monto(valor_usuario, valor_doc):
    u = normalizar_monto(valor_usuario)
    d = normalizar_monto(valor_doc)
    if u is None:
        return {"estado": "SIN_DATO", "encontrado": valor_doc, "mensaje": None}
    if d is None:
        return {"estado": "NO_ENCONTRADO", "encontrado": None, "mensaje": "No se pudo leer el monto del documento"}
    if u == d:
        return {"estado": "CORRECTO", "encontrado": valor_doc, "mensaje": None}
    return {
        "estado": "DIFERENTE",
        "encontrado": valor_doc,
        "mensaje": f'Se ingresó {formatear_monto(u)} pero el documento tiene {formatear_monto(d)}',
    }


def _comparar_fecha(valor_usuario, valor_doc):
    if not valor_usuario:
        return {"estado": "SIN_DATO", "encontrado": valor_doc, "mensaje": None}
    if not valor_doc:
        return {"estado": "NO_ENCONTRADO", "encontrado": None, "mensaje": "No se encontró la fecha en el documento"}
    u_str = str(valor_usuario)
    u_norm = u_str[:10] if len(u_str) >= 10 else u_str
    if u_norm == valor_doc:
        return {"estado": "CORRECTO", "encontrado": valor_doc, "mensaje": None}
    return {
        "estado": "DIFERENTE",
        "encontrado": valor_doc,
        "mensaje": f'Se ingresó "{valor_usuario}" pero el documento tiene "{valor_doc}"',
    }


def _comparar_contacto(valor_usuario, valor_doc):
    u = re.sub(r'\D', '', str(valor_usuario))
    d = re.sub(r'\D', '', str(valor_doc))
    if u == d:
        return {"estado": "CORRECTO", "encontrado": valor_doc, "mensaje": None}
    return {
        "estado": "DIFERENTE",
        "encontrado": valor_doc,
        "mensaje": f'Se ingresó "{valor_usuario}" pero el documento tiene "{valor_doc}"',
    }


# Mapeo tipo_campo → comparador directo
_COMPARADORES_DIRECTOS = {
    "nombre":    _comparar_nombre,
    "texto":     _comparar_texto,
    "monto":     _comparar_monto,
    "fecha":     _comparar_fecha,
    "contacto":  _comparar_contacto,
}


# =========================================================================
# VALIDACIÓN CON EXTRACCIÓN ESTRUCTURADA (campos adicionales)
# =========================================================================

def validar_campos_con_extraccion_estructurada(tipo_documento, form_data, campos_doc):
    """
    Compara campos específicos (cargo, salario, etc.) contra los extraídos.
    NOTA: nombre y documento se validan por separado.
    """
    claves_aplicables = CAMPOS_POR_TIPO.get(tipo_documento, [])
    resultados = []

    for clave in claves_aplicables:
        defn = DEFINICIONES_CAMPOS.get(clave, {})
        valor_usuario = form_data.get(clave)

        if not valor_usuario:
            continue

        if clave in ("nombre", "documento"):
            continue

        label    = defn.get("label", clave)
        tipo_cmp = defn.get("tipo", "texto")
        valor_doc = campos_doc.get(clave)

        if valor_doc is None:
            resultado = {
                "estado": "NO_ENCONTRADO",
                "encontrado": None,
                "mensaje": "No se encontró este campo en el documento",
            }
        else:
            comparador = _COMPARADORES_DIRECTOS.get(tipo_cmp, _comparar_texto)
            resultado  = comparador(valor_usuario, valor_doc)

        resultados.append({
            "label":         label,
            "valor_usuario": str(valor_usuario),
            "tipo":          tipo_cmp,
            **resultado,
        })

    return resultados


# =========================================================================
# ESTADOS INCORRECTOS Y FUNCIÓN PARA DETERMINAR ESTADO GENERAL
# =========================================================================

ESTADOS_INCORRECTOS = {
    "DIFERENTE",
    "DOCUMENTO_INCORRECTO",
    "NO_ENCONTRADO",
    "INCONSISTENCIA",
    "PARCIAL",
}


def _estado_es_incorrecto(estado):
    return estado in ESTADOS_INCORRECTOS


def determinar_estado_general(resultado_nombre, resultado_documento, resultado_campos):
    """
    Determina el estado GENERAL del archivo, combinando:
      - la validación de nombre
      - la validación de documento (cédula)
      - todos los campos adicionales (cargo, salario, auxilios, etc.)

    Si CUALQUIERA de estas validaciones resulta incorrecta, el archivo
    completo se marca "INCORRECTO" y hay que revisarlo. Solo si TODO
    está correcto (o los campos vacíos/omitidos no aplican) se marca
    "CORRECTO".
    """
    hay_error = False

    # --- Nombre ---
    estado_nombre_directo = resultado_nombre.get("_estado_directo", {}).get("estado")
    if estado_nombre_directo:
        if _estado_es_incorrecto(estado_nombre_directo):
            hay_error = True
    else:
        if resultado_nombre.get("total_ocurrencias", 0) == 0 or resultado_nombre.get("con_errores", 0) > 0:
            hay_error = True

    # --- Documento (cédula) ---
    estado_documento = resultado_documento.get("estado")
    if estado_documento and estado_documento != "OMITIDO" and _estado_es_incorrecto(estado_documento):
        hay_error = True

    # --- Campos adicionales ---
    for campo in resultado_campos:
        if _estado_es_incorrecto(campo.get("estado")):
            hay_error = True
            break

    if hay_error:
        return {
            "estado": "INCORRECTO",
            "mensaje": "❌ Documento Incorrecto: valida la información.",
        }

    return {
        "estado": "CORRECTO",
        "mensaje": "✅ Documento Correcto.",
    }


# =========================================================================
# VISTA PRINCIPAL (MEJORADA + SESIÓN)
# =========================================================================

def validar_documentos(request):
    # --- Limpiar sesión si se pide explícitamente ---
    if request.GET.get("limpiar") == "1":
        request.session.pop("form_data", None)
        return redirect("validar_documentos")

    # --- Recuperar datos guardados en sesión (si existen) ---
    form_data_sesion = request.session.get("form_data", {})

    resultados     = []
    error_archivos = None
    datos_para_mostrar = form_data_sesion  # valor por defecto (GET)

    if request.method == "POST":
        # Guardar todos los datos del POST en la sesión, preservando valores anteriores si el nuevo está vacío
        nuevos_datos = {}
        for campo in CAMPOS_FORMULARIO:
            valor_nuevo = request.POST.get(campo, "").strip()
            if valor_nuevo:
                nuevos_datos[campo] = valor_nuevo
            else:
                nuevos_datos[campo] = form_data_sesion.get(campo, "")
        request.session["form_data"] = nuevos_datos
        request.session.modified = True
        datos_para_mostrar = nuevos_datos  # usar los datos ya combinados/actualizados

        form = ValidacionForm(request.POST)

        if form.is_valid():
            nombre    = nuevos_datos["nombre"]
            documento = nuevos_datos["documento"]

            archivos      = request.FILES.getlist("archivos")
            tipos_archivo = request.POST.getlist("tipo_archivo[]")

            if not archivos:
                error_archivos = "Debes subir al menos un archivo."
            else:
                for idx, archivo in enumerate(archivos):
                    tipo_archivo = tipos_archivo[idx] if idx < len(tipos_archivo) else None

                    if not tipo_archivo or tipo_archivo not in TIPOS_DOCUMENTOS:
                        error_archivos = f"El archivo '{archivo.name}' no tiene un tipo válido asignado."
                        continue

                    # 1. Extraer campos estructurados por tipo
                    campos_doc = extraer_campos_estructurados(tipo_archivo, archivo)
                    archivo.seek(0)

                    # 2. Extraer texto plano
                    texto = extraer_texto(archivo)

                    # 3. NOMBRE: usa extracción estructurada si existe, sino texto plano
                    if campos_doc.get("nombre"):
                        resultado_nombre_campo = _comparar_nombre(nombre, campos_doc["nombre"])
                        resultado_nombre = {
                            "nombre_usuario": nombre,
                            "total_ocurrencias": 1,
                            "exactas": 1 if resultado_nombre_campo["estado"] == "CORRECTO" else 0,
                            "con_errores": 0 if resultado_nombre_campo["estado"] == "CORRECTO" else 1,
                            "porcentaje_promedio": 100 if resultado_nombre_campo["estado"] == "CORRECTO" else 0,
                            "ocurrencias": [{
                                "texto": campos_doc["nombre"],
                                "score": 100 if resultado_nombre_campo["estado"] == "CORRECTO" else 0,
                                "es_exacta": resultado_nombre_campo["estado"] == "CORRECTO",
                                "detalles": [],
                                "adicionales": [],
                            }],
                            "_estado_directo": resultado_nombre_campo,
                        }
                    else:
                        resultado_nombre = validar_nombre(nombre, texto)

                    # 4. DOCUMENTO: se valida solo si NO es CONDICIONES ni MEDICAS
                    if tipo_archivo in ("CONDICIONES", "MEDICAS"):
                        resultado_documento = {
                            "documento_usuario": documento,
                            "tipo_documento": "CC",
                            "descripcion_tipo": "No aplica",
                            "total_ocurrencias": 0,
                            "exactas": 0,
                            "con_errores": 0,
                            "inconsistencia_entre_ocurrencias": False,
                            "estado": "OMITIDO",
                            "ocurrencias": [],
                            "mensaje": "Validación de documento omitida para este tipo de documento (CONDICIONES o MEDICAS).",
                            "todos_numeros_documento": [],
                            "validacion_nit": None,
                            "contextos_cc": [],
                            "_estado_directo": {
                                "estado": "OMITIDO",
                                "mensaje": "Omitido por tipo CONDICIONES o MEDICAS"
                            },
                        }
                    else:
                        resultado_documento_mejorado = extraer_y_validar_documento(
                            documento,
                            texto,
                            campos_doc_estructurados=campos_doc,
                        )

                        estado_map = {
                            "CORRECTO": "CORRECTO",
                            "DIFERENTE": "DOCUMENTO_INCORRECTO",
                            "NO_ENCONTRADO": "NO_ENCONTRADO",
                            "INCONSISTENCIA": "INCONSISTENCIA",
                        }

                        resultado_documento = {
                            "documento_usuario": documento,
                            "tipo_documento": "CC",
                            "descripcion_tipo": "Extraído del documento",
                            "total_ocurrencias": len(resultado_documento_mejorado.get("contextos", [])),
                            "exactas": 1 if resultado_documento_mejorado["estado"] == "CORRECTO" else 0,
                            "con_errores": (1 if resultado_documento_mejorado["estado"] in ("DIFERENTE", "INCONSISTENCIA") else 0),
                            "inconsistencia_entre_ocurrencias": resultado_documento_mejorado["estado"] == "INCONSISTENCIA",
                            "estado": estado_map.get(resultado_documento_mejorado["estado"], "DOCUMENTO_INCORRECTO"),
                            "ocurrencias": [
                                {
                                    "numero": ctx["numero_limpio"],
                                    "similitud": 100 if ctx["numero_limpio"] == re.sub(r'\D', '', documento) else 0,
                                    "es_exacto": ctx["numero_limpio"] == re.sub(r'\D', '', documento),
                                    "largo": len(ctx["numero_limpio"]),
                                    "diferencia_detalle": None,
                                }
                                for ctx in resultado_documento_mejorado.get("contextos", [])
                            ],
                            "mensaje": resultado_documento_mejorado["mensaje"],
                            "todos_numeros_documento": [
                                ctx["numero_limpio"] for ctx in resultado_documento_mejorado.get("contextos", [])
                            ],
                            "validacion_nit": None,
                            "contextos_cc": [
                                {
                                    "etiqueta": ctx.get("etiqueta", "C.C."),
                                    "numero_encontrado": ctx["numero_limpio"],
                                    "antes": "",
                                    "despues": "",
                                    "coincide": ctx["numero_limpio"] == re.sub(r'\D', '', documento),
                                    "diferencia": None if ctx["numero_limpio"] == re.sub(r'\D', '', documento) else (
                                        f"Se ingresó '{documento}' pero encontrado '{ctx['numero_limpio']}'"
                                    ),
                                }
                                for ctx in resultado_documento_mejorado.get("contextos", [])
                            ],
                            "_estado_directo": resultado_documento_mejorado,
                        }

                    # 5. CAMPOS ADICIONALES
                    resultado_campos = validar_campos_con_extraccion_estructurada(
                        tipo_archivo,
                        nuevos_datos,
                        campos_doc,
                    )

                    # 6. Determinar estado general combinando nombre, documento y campos
                    estado_general = determinar_estado_general(
                        resultado_nombre, resultado_documento, resultado_campos
                    )

                    resultados.append({
                        "archivo":      archivo.name,
                        "tipo_archivo": TIPOS_DOCUMENTOS.get(tipo_archivo, tipo_archivo),
                        "nombre":       resultado_nombre,
                        "documento":    resultado_documento,
                        "campos":       resultado_campos,
                        "tipo_riesgo":  campos_doc.get("tipo_riesgo"),
                        "estado_general": estado_general,
                    })
        else:
            print("ERRORES FORM:", form.errors)
    else:
        form = ValidacionForm(initial=form_data_sesion)

    initial_data_json = json.dumps(datos_para_mostrar, ensure_ascii=False)

    return render(
        request,
        "validar_documentos.html",
        {
            "form":             form,
            "resultados":       resultados,
            "error_archivos":   error_archivos,
            "tipos_documentos": TIPOS_DOCUMENTOS,
            "initial_data_json": initial_data_json,
        },
    )