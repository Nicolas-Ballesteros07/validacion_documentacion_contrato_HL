"""
buscar_documento_mejorado.py
----------------------------
Busca agresivamente TODOS los números de CC/Documento en el texto
y reporta claramente qué se encontró vs. qué se ingresó.

Se aplica cuando:
1. La extracción estructurada NO encontró documento
2. O cuando queremos verificar consistencia completa
"""

import re
from rapidfuzz import fuzz


# Todos los patrones posibles de etiquetas de documento
PATRONES_DOCUMENTO = [
    # Variantes de C.C
    r'C\.?\s*C\.?',
    r'C[eéEÉ]dula\s+de\s+ciudadan[ií]a',
    r'C[eéEÉ]dula\s+de\s+identidad',
    r'Documento\s+de\s+identidad',
    r'Número\s+de\s+documento',
    r'Cédula',
    r'C\.?I\.?',  # Cédula de Identidad
]

# Compilar patrón único que busca etiqueta + número
PATRON_DOCUMENTO_COMPLETO = re.compile(
    r'(?:' + '|'.join(PATRONES_DOCUMENTO) + r')\s*[:\-]?\s*(\d[\d.\s]{4,}\d|\d{5,})',
    re.IGNORECASE
)


def limpiar_numero(texto):
    """Extrae solo dígitos de un número."""
    return re.sub(r'\D', '', str(texto))


def buscar_todos_documentos_en_texto(texto):
    """
    Busca TODOS los números de documento/CC en el texto.

    Returns:
        Lista de dicts con:
        {
            "numero_limpio": "73144545",
            "numero_original": "C.C. 73144545 Expedida en CARTAGENA",
            "etiqueta": "C.C.",
            "linea": "C.C. 73144545 Expedida en CARTAGENA",
            "posicion": 1234,
        }
    """
    encontrados = []

    for match in PATRON_DOCUMENTO_COMPLETO.finditer(texto):
        numero_raw = match.group(1)
        numero_limpio = limpiar_numero(numero_raw)

        # Filtrar números muy cortos (menos de 5 dígitos) o muy largos (más de 15)
        if len(numero_limpio) < 5 or len(numero_limpio) > 15:
            continue

        # Obtener la línea completa donde aparece
        inicio_linea = max(0, match.start() - 100)
        fin_linea = min(len(texto), match.end() + 100)
        contexto = texto[inicio_linea:fin_linea].strip()

        # Extraer solo la línea actual (hasta primer salto de línea)
        linea_actual = contexto.split('\n')[0]

        encontrados.append({
            "numero_limpio": numero_limpio,
            "numero_original": numero_raw.strip(),
            "etiqueta": match.group(0)[:30],  # Primeros 30 chars del match
            "linea": linea_actual,
            "posicion": match.start(),
        })

    return encontrados


def validar_documento_mejorado(documento_usuario, texto):
    """
    Valida documento buscando agresivamente en el texto.

    Si NO encuentra contextos de CC, devuelve NO_ENCONTRADO.
    Si encuentra contextos pero NINGUNO coincide, devuelve ERROR.
    Si encuentra múltiples números DIFERENTES, devuelve INCONSISTENCIA.

    Returns:
        {
            "estado": "CORRECTO" | "DIFERENTE" | "NO_ENCONTRADO" | "INCONSISTENCIA",
            "encontrado": "número encontrado o None",
            "contextos": [list de contextos encontrados],
            "mensaje": "mensaje descriptivo",
            "documento_usuario": documento ingresado,
        }
    """

    doc_usuario_limpio = limpiar_numero(documento_usuario)
    contextos_encontrados = buscar_todos_documentos_en_texto(texto)

    # NO ENCONTRADO: si no hay contextos de CC
    if not contextos_encontrados:
        return {
            "estado": "NO_ENCONTRADO",
            "encontrado": None,
            "contextos": [],
            "mensaje": (
                f"❌ CÉDULA NO ENCONTRADA: No se encontraron etiquetas de "
                f"C.C., Cédula o Documento en el archivo. Se buscó '{documento_usuario}'."
            ),
            "documento_usuario": documento_usuario,
        }

    # Verificar si alguno coincide exactamente
    coincidencias_exactas = [
        ctx for ctx in contextos_encontrados
        if ctx["numero_limpio"] == doc_usuario_limpio
    ]

    # CORRECTO: si hay coincidencia exacta
    if coincidencias_exactas:
        return {
            "estado": "CORRECTO",
            "encontrado": coincidencias_exactas[0]["numero_limpio"],
            "contextos": contextos_encontrados,
            "mensaje": (
                f"✅ CÉDULA CORRECTA: '{documento_usuario}' encontrado en "
                f"{len(coincidencias_exactas)} ocurrencia(s) con etiquetas de C.C./Documento."
            ),
            "documento_usuario": documento_usuario,
        }

    # INCONSISTENCIA o DIFERENTE: si encontró contextos pero no coinciden
    números_únicos = set(ctx["numero_limpio"] for ctx in contextos_encontrados)

    if len(números_únicos) > 1:
        # Múltiples números diferentes en el documento
        numeros_str = ", ".join(sorted(números_únicos))
        return {
            "estado": "INCONSISTENCIA",
            "encontrado": list(números_únicos)[0],  # Primer número encontrado
            "contextos": contextos_encontrados,
            "mensaje": (
                f"⚠️ INCONSISTENCIA EN EL DOCUMENTO: Se ingresó '{documento_usuario}' "
                f"pero se encontraron MÚLTIPLES números de C.C./Documento DIFERENTES: {numeros_str}. "
                f"El documento tiene inconsistencias internas que deben verificarse."
            ),
            "documento_usuario": documento_usuario,
        }
    else:
        # Un único número pero no coincide
        numero_encontrado = list(números_únicos)[0]
        return {
            "estado": "DIFERENTE",
            "encontrado": numero_encontrado,
            "contextos": contextos_encontrados,
            "mensaje": (
                f"❌ DOCUMENTO INCORRECTO: Se ingresó '{documento_usuario}' "
                f"pero el documento tiene '{numero_encontrado}'. "
                f"Se encontraron {len(contextos_encontrados)} contexto(s) de C.C./Documento."
            ),
            "documento_usuario": documento_usuario,
        }


# ============================================================================
# Función para integrar con views.py
# ============================================================================

def extraer_y_validar_documento(documento_usuario, texto, campos_doc_estructurados=None):
    """
    Inteligencia completa para validar documento:
    1. Si hay extracción estructurada → úsala primero
    2. Si no hay → busca agresivamente en texto
    3. Reporta claramente qué encontró

    Args:
        documento_usuario: str con el documento ingresado por el usuario
        texto: str con el texto extraído del PDF/DOCX
        campos_doc_estructurados: dict opcional con campos extraídos estructuradamente

    Returns:
        {
            "estado": "CORRECTO" | "DIFERENTE" | "NO_ENCONTRADO" | "INCONSISTENCIA",
            "encontrado": valor encontrado,
            "contextos": lista de contextos,
            "mensaje": mensaje claro,
            "fuente": "estructurado" | "búsqueda_texto",
        }
    """

    doc_usuario_limpio = limpiar_numero(documento_usuario)

    # Intentar con extracción estructurada primero
    if campos_doc_estructurados and campos_doc_estructurados.get("documento"):
        doc_extraído = campos_doc_estructurados["documento"]
        doc_extraído_limpio = limpiar_numero(doc_extraído)
        es_exacto = doc_usuario_limpio == doc_extraído_limpio

        # ✅ NUEVO: aunque el estructurado coincida, buscar TODOS los números
        # de CC en el texto para detectar inconsistencias internas
        contextos_texto = buscar_todos_documentos_en_texto(texto)
        numeros_distintos = set(
            ctx["numero_limpio"] for ctx in contextos_texto
            if ctx["numero_limpio"] != doc_usuario_limpio
        )

        # Si el estructurado no coincide con el usuario → error directo
        if not es_exacto:
            return {
                "estado": "DIFERENTE",
                "encontrado": doc_extraído,
                "contextos": [{
                    "numero_limpio": doc_extraído_limpio,
                    "numero_original": doc_extraído,
                    "fuente": "extracción estructurada",
                    "linea": f"Extraído de tabla: {doc_extraído}",
                    "posicion": -1,
                }],
                "mensaje": (
                    f"❌ CÉDULA INCORRECTA: Se ingresó '{documento_usuario}' "
                    f"pero la extracción estructurada tiene '{doc_extraído}'"
                ),
                "documento_usuario": documento_usuario,
                "fuente": "estructurado",
            }

        # Si coincide pero hay otros números de CC diferentes en el texto → inconsistencia
        if numeros_distintos:
            numeros_str = ", ".join(sorted(numeros_distintos))
            return {
                "estado": "INCONSISTENCIA",
                "encontrado": doc_extraído,
                "contextos": contextos_texto,
                "mensaje": (
                    f"⚠️ INCONSISTENCIA EN EL DOCUMENTO: '{documento_usuario}' coincide en las tablas "
                    f"de firma, pero se encontraron otros números de C.C. DIFERENTES en el texto: "
                    f"{numeros_str}. Verificar el documento completo."
                ),
                "documento_usuario": documento_usuario,
                "fuente": "estructurado+texto",
            }

        # Todo correcto
        return {
            "estado": "CORRECTO",
            "encontrado": doc_extraído,
            "contextos": [{
                "numero_limpio": doc_extraído_limpio,
                "numero_original": doc_extraído,
                "fuente": "extracción estructurada",
                "linea": f"Extraído de tabla: {doc_extraído}",
                "posicion": -1,
            }],
            "mensaje": (
                f"✅ CÉDULA CORRECTA: '{documento_usuario}' (estructurado)"
            ),
            "documento_usuario": documento_usuario,
            "fuente": "estructurado",
        }

    # Si no hay extracción estructurada, buscar en texto
    return {
        **validar_documento_mejorado(documento_usuario, texto),
        "fuente": "búsqueda_texto",
    }