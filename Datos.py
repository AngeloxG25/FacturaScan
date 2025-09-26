# =========================
# Datos de las empresas
# =========================
RAZONES_TXT = r"""
COMERCIAL TEBA SPA;76.466.343-8;Gran Avenida=Avenida Jose Miguel Carrera #13365, San Bernardo|Lo Valledor=Avenida General Velázquez #3409, Cerrillos|Rancagua=Avenida Federico Koke #250, Rancagua|JJ Perez=Avenida Jose Joaquin Perez #6142, Cerro Navia|Oficina Central=Avenida Lo Blanco #2561, La Pintana|Pinto=Pinto 8, San Bernardo|Lo Blanco=Avenida Lo Blanco #2561, La Pintana
COMERCIAL NABEK LIMITADA;78.767.200-0;Nabek Lo Blanco=Avenida Lo Blanco #2561, La Pintana|Oficina Central=Avenida Lo Blanco #2561, La Pintana
RECURSOS HUMANOS A TIEMPO SPA;77.076.847-0;Lo Blanco=Avenida Lo Blanco #2561, La Pintana|Oficina Central=Avenida Lo Blanco #2561, La Pintana
TRANSPORTE LS SPA;76.704.181-0;Lo Blanco=Avenida Lo Blanco #2561, La Pintana|Oficina Central=Avenida Lo Blanco #2561, La Pintana
INMOBILIARIA NABEK SPA;77.963.143-5;Lo Blanco=Avenida Lo Blanco #2561, La Pintana|Oficina Central=Avenida Lo Blanco #2561, La Pintana
TEAM WORK NABEK SPA;78.075.668-3;Lo Blanco=Avenida Lo Blanco #2561, La Pintana|Oficina Central=Avenida Lo Blanco #2561, La Pintana
""".strip()

# =========================
# 2) Carpetas OneDrive candidatas para CONTROL_DOCUMENTAL
# (usa los nombres EXACTOS tal como aparecen en tu OneDrive)
# =========================
CONTROL_DOC_CANDIDATES = [
    "Archivos de Documentos Teba - CONTROL_DOCUMENTAL",
    "CONTROL_DOCUMENTAL",
]

# =========================
# 3) Mapeo Razón Social -> carpeta empresa en CONTROL_DOCUMENTAL
# (keys en minúscula, sin tildes; values = nombre EXACTO de carpeta)
# Debe coincidir con tu estructura en OneDrive:
#   INMOBILIARIA_NABEK, NABEK, RRHH_A_TIEMPO, TEAMWORK, TEBA, TRANSPORTES_LS
# =========================
COMPANY_ROOT_BY_RAZON = {
    "comercial teba spa":             "TEBA",
    "comercial nabek limitada":       "NABEK",
    "recursos humanos a tiempo spa":  "RRHH_A_TIEMPO",
    "transporte ls spa":              "TRANSPORTES_LS",
    "inmobiliaria nabek spa":         "INMOBILIARIA_NABEK",
    "team work nabek spa":            "TEAMWORK",
}

# =========================
# 4) Códigos de sucursal por empresa (nombres de carpetas de sucursal)
#    - TEBA usa sus códigos numéricos oficiales
#    - Para el resto, se define 'oficina central' -> 005_OFICINA_CENTRAL
# =========================
SUC_CODE_BY_COMPANY = {
    "TEBA": {
        "gran avenida":    "001_GRAN_AVENIDA",
        "lo valledor":     "002_LO_VALLEDOR",
        "rancagua":        "003_RANCAGUA",
        "jj perez":        "004_JJ_PEREZ",
        "oficina central": "005_OFICINA_CENTRAL",
        "pinto":           "007_PINTO",
        "lo blanco":       "009_LO_BLANCO",
    },
    "NABEK": {
        "nabek lo blanco": "103_LO_BLANCO",
        "lo blanco":       "103_LO_BLANCO",
        "oficina central": "005_OFICINA_CENTRAL",
    },
    "RRHH_A_TIEMPO": {
        "lo blanco":       "LO_BLANCO",
        "oficina central": "005_OFICINA_CENTRAL",
    },
    "TRANSPORTES_LS": {
        "lo blanco":       "LO_BLANCO",
        "oficina central": "005_OFICINA_CENTRAL",
    },
    "INMOBILIARIA_NABEK": {
        "lo blanco":       "LO_BLANCO",
        "oficina central": "005_OFICINA_CENTRAL",
    },
    "TEAMWORK": {
        "lo blanco":       "LO_BLANCO",
        "oficina central": "005_OFICINA_CENTRAL",
    },
}
