
# Datos de las empresas
RAZONES_TXT = r"""
COMERCIAL TEBA SPA;76.466.343-8;Gran Avenida=Avenida Jose Miguel Carrera #13365, San Bernardo|Lo Valledor=Avenida General Velázquez #3409, Cerrillos|Rancagua=Avenida Federico Koke #250, Rancagua|JJ Perez=Avenida Jose Joaquin Perez #6142, Cerro Navia|Pinto=Pinto 8, San Bernardo|Lo Blanco=Avenida Lo Blanco #2561, La Pintana
COMERCIAL NABEK LIMITADA;78.767.200-0;Nabek Lo Blanco=Avenida Lo Blanco #2561, La Pintana
""".strip()

# RESPALDO

# RAZONES_TXT = r"""
# COMERCIAL TEBA SPA;76.466.343-8;Gran Avenida=Avenida Jose Miguel Carrera #13365, San Bernardo|Lo Valledor=Avenida General Velázquez #3409, Cerrillos|Rancagua=Avenida Federico Koke #250, Rancagua|JJ Perez=Avenida Jose Joaquin Perez #6142, Cerro Navia|Pinto=Pinto 8, San Bernardo|Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# COMERCIAL NABEK LIMITADA;78.767.200-0;Nabek Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# RECURSOS HUMANOS A TIEMPO SPA;77.076.847-0;Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# TRANSPORTE LS SPA;76.704.181-0;Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# INMOVILIARIA NABEK SPA;77.963.143-5;Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# TEAM WORK NABEK SPA;78.075.668-3;Lo Blanco=Avenida Lo Blanco #2561, La Pintana
# JUANITO;1-9;juanito=juanito 123, san miguel
# """.strip()

# --- 2) Carpetas OneDrive donde puede estar CONTROL_DOCUMENTAL ---
# (pon aquí el/los nombres EXACTOS que ves en tu OneDrive)
CONTROL_DOC_CANDIDATES = [
    "Archivos de Documentos Teba - CONTROL_DOCUMENTAL", 
    "CONTROL_DOCUMENTAL",                               
]

# --- 3) Mapeo de nombre de razón -> carpeta empresa en CONTROL_DOCUMENTAL ---
# (keys en minúscula, sin tildes; values = nombre de carpeta)
COMPANY_ROOT_BY_RAZON = {
    "comercial teba spa": "TEBA",
    "comercial nabek limitada": "NABEK",
    # agrega otros si quieres que apunten a carpetas propias:
    "recursos humanos a tiempo spa": "RHT",
    "transporte ls spa": "TRANSPORTE_LS",
    "inmobiliaria nabek spa": "INMOBILIARIA_NABEK",
    "team work nabek spa": "TEAM_WORK_NABEK",
    "juanito": "JUANITO",
}

# --- 4) Códigos de sucursal por empresa (como aparece en OneDrive) ---
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
    },
    # se puede agrega más empresas si corresponde…
}

