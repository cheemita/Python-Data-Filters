import pandas as pd
import re  # Librería para expresiones regulares

# Lista de códigos de área válidos (sólo EE. UU. y Canadá)
CODIGOS_AREA_VALIDOS = {
    "201", "202", "203", "204", "205", "206", "207", "208", "209", "210", "212", "213", "214", "215", "216", "217", "218", 
    "219", "224", "225", "226", "228", "229", "231", "234", "239", "240", "242", "246", "248", "250", "251", "252", "253",
    "254", "256", "260", "262", "264", "267", "268", "269", "270", "272", "276", "279", "281", "283", "289", "301", "302",
    "303", "304", "305", "306", "307", "308", "309", "310", "312", "313", "314", "315", "316", "317", "318", "319", "320",
    "321", "323", "325", "326", "330", "331", "334", "336", "337", "339", "340", "343", "345", "346", "347", "351", "352",
    "360", "361", "364", "365", "380", "385", "386", "401", "402", "403", "404", "405", "406", "407", "408", "409", "410",
    "412", "413", "414", "415", "416", "417", "418", "419", "423", "424", "425", "430", "431", "432", "434", "435", "437",
    "438", "440", "441", "442", "443", "450", "456", "458", "463", "469", "470", "473", "475", "478", "479", "480", "481",
    "484", "501", "502", "503", "504", "505", "506", "507", "508", "509", "510", "512", "513", "514", "515", "516", "517",
    "518", "519", "520", "530", "531", "533", "534", "539", "540", "541", "548", "551", "557", "559", "561", "562", "563",
    "564", "567", "570", "571", "573", "574", "575", "579", "580", "581", "582", "585", "586", "601", "602", "603", "604",
    "605", "606", "607", "608", "609", "610", "612", "613", "614", "615", "616", "617", "618", "619", "620", "623", "626",
    "627", "628", "629", "630", "631", "636", "639", "640", "641", "646", "647", "649", "650", "651", "657", "658", "659",
    "660", "661", "662", "664", "667", "669", "670", "671", "678", "679", "680", "681", "682", "684", "689", "701", "702",
    "703", "704", "705", "706", "707", "708", "709", "712", "713", "714", "715", "716", "717", "718", "719", "720", "721",
    "724", "725", "726", "727", "730", "731", "732", "734", "737", "740", "743", "747", "754", "757", "758", "760", "762",
    "763", "764", "765", "767", "769", "770", "771", "772", "773", "774", "775", "778", "779", "780", "781", "782", "784",
    "785", "786", "787", "801", "802", "803", "804", "805", "806", "807", "808", "809", "810", "812", "813", "814", "815",
    "816", "817", "818", "819", "820", "825", "828", "829", "830", "831", "832", "835", "843", "844", "845", "847", "848",
    "849", "850", "854", "855", "856", "857", "858", "859", "860", "862", "863", "864", "865", "867", "868", "869", "870",
    "872", "873", "876", "878", "901", "902", "903", "904", "905", "906", "907", "908", "909", "910", "912", "913", "914",
    "915", "916", "917", "918", "919", "920", "925", "927", "928", "929", "930", "931", "935", "936", "937", "938", "939",
    "940", "941", "947", "949", "951", "952", "954", "956", "959", "970", "971", "972", "973", "975", "978", "979", "980",
    "984", "985", "986", "989"
}

# Función para leer el archivo Excel
def leer_archivo_excel(ruta_archivo):
    try:
        df = pd.read_excel(ruta_archivo, header=None)  # Leer sin encabezados por defecto
        encabezados = [f"Columna_{i + 1}" for i in range(df.shape[1])]  # Encabezados genéricos
        df.columns = encabezados
        return encabezados, df
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        exit()

# Función para mostrar encabezados
def mostrar_encabezados(encabezados):
    print("\nEncabezados del archivo:")
    for i, encabezado in enumerate(encabezados):
        print(f"{i + 1}. {encabezado}")
    print("\nSeleccione los números de las columnas que desea mantener, separados por comas.")
    print("Para finalizar y mantener las columnas seleccionadas, presione 0.")

# Función para filtrar columnas seleccionadas
def filtrar_datos_para_mantener(df, columnas_a_mantener):
    return df[columnas_a_mantener]

# Función para validar números telefónicos
def validar_telefonos(df):
    """
    Valida y limpia números telefónicos en las columnas pertinentes.
    Detecta automáticamente columnas que contienen números de teléfono,
    y elimina registros con números inválidos.
    """
    for columna in df.columns:
        # Detectar si la columna contiene números telefónicos analizando una muestra
        muestra = df[columna].astype(str).head(10).tolist()

        def es_posible_telefono(valor):
            """
            Comprueba si un valor podría ser un número de teléfono basado en su longitud
            y formato esperado (10 dígitos y un código de área válido).
            """
            valor = re.sub(r'\D', '', valor)  # Eliminar caracteres no numéricos
            return len(valor) == 10 and valor[:3] in CODIGOS_AREA_VALIDOS

        # Determinar si más de la mitad de la muestra podría ser números de teléfono
        es_columna_telefonica = sum([es_posible_telefono(valor) for valor in muestra]) > (len(muestra) // 2)

        if es_columna_telefonica:
            print(f"Detectando y validando números telefónicos en la columna '{columna}'...")

            # Convertir todos los valores a string
            df[columna] = df[columna].astype(str)

            def limpiar_y_validar(numero):
                """
                Limpia y valida un número telefónico.
                Retorna el número limpio si es válido, o None si no lo es.
                """
                numero = re.sub(r'\D', '', numero)  # Eliminar caracteres no numéricos
                if len(numero) != 10:  # Debe tener exactamente 10 dígitos
                    return None
                if numero[:3] not in CODIGOS_AREA_VALIDOS:  # Validar código de área
                    return None
                return numero  # Retornar el número limpio si es válido

            # Limpiar y validar la columna
            df[f"{columna}_limpio"] = df[columna].apply(limpiar_y_validar)

            # Mostrar estadísticas
            total_validos = df[f"{columna}_limpio"].notna().sum()
            total_invalidos = len(df) - total_validos
            print(f" - Total válidos: {total_validos}")
            print(f" - Total inválidos: {total_invalidos}")

            # Filtrar registros válidos y reemplazar la columna original por los números limpios
            df = df[df[f"{columna}_limpio"].notna()]
            df[columna] = df[f"{columna}_limpio"]
            df = df.drop(columns=[f"{columna}_limpio"])

    return df

# Función para concatenar nombres y apellidos
def concatenar_nombres_apellidos(df, col_nombre, col_apellido):
    """
    Concatena las columnas de nombre y apellido en una nueva columna 'Nombre Completo'.
    Si ambos campos están vacíos, elimina el registro.
    """
    # Sustituir NaN por espacio vacío
    df[col_nombre] = df[col_nombre].fillna("")
    df[col_apellido] = df[col_apellido].fillna("")

    # Crear la nueva columna 'Nombre Completo'
    df["name"] = df[col_nombre].astype(str) + " " + df[col_apellido].astype(str)

    # Eliminar registros donde ambas columnas sean vacías
    df = df[~((df[col_nombre] == "") & (df[col_apellido] == ""))]

    return df

# Función para procesar nombres en una sola columna
def procesar_nombre_completo(df, col_nombre_completo):
    try:
        df["name"] = df[col_nombre_completo].astype(str)
        df.drop(columns=[col_nombre_completo], inplace=True)
        return df
    except Exception as e:
        print(f"Error al procesar el nombre completo: {e}")
        return df

# Función principal
def limpiar_archivo_excel(ruta_archivo):
    """
    Proceso principal para leer, limpiar y guardar el archivo Excel.
    """
    # Leer el archivo y generar encabezados iniciales
    encabezados, df = leer_archivo_excel(ruta_archivo)

    # Mostrar encabezados al usuario
    mostrar_encabezados(encabezados)

    print("\n¿Desea procesar nombres y apellidos?")
    print("1. Concatenar dos columnas (nombre y apellido).")
    print("2. Usar una sola columna para el nombre completo.")
    opcion = input("Seleccione una opción (1 o 2): ").strip()

    if opcion == '1':
        col_nombre = input("Ingrese el número de la columna para 'Nombre': ").strip()
        col_apellido = input("Ingrese el número de la columna para 'Apellido': ").strip()
        try:
            col_nombre = encabezados[int(col_nombre) - 1]
            col_apellido = encabezados[int(col_apellido) - 1]
            df = concatenar_nombres_apellidos(df, col_nombre, col_apellido)
            encabezados = df.columns.tolist()  # Actualizar encabezados
            print("\nColumna 'Nombre Completo' creada y agregada. Registros vacíos eliminados.")
        except (ValueError, IndexError):
            print("Error al seleccionar columnas para concatenar.")

    elif opcion == '2':
        col_nombre_completo = input("Ingrese el número de la columna para 'Name': ").strip()
        try:
            col_nombre_completo = encabezados[int(col_nombre_completo) - 1]
            df = procesar_nombre_completo(df, col_nombre_completo)
            encabezados = df.columns.tolist()  # Actualizar encabezados
            print("\nColumna 'name' procesada y agregada.")
        except (ValueError, IndexError):
            print("Error al seleccionar columna para nombre completo.")

    # Reorganizar la columna 'Nombre Completo' al inicio si existe
    if 'name' in encabezados:
        cols = ['name'] + [col for col in encabezados if col != 'name']
        df = df[cols]
        encabezados = df.columns.tolist()

    # Seleccionar columnas adicionales
    columnas_a_mantener = []
    while True:
        mostrar_encabezados(encabezados)
        seleccion = input("\nIngrese los números de las columnas que desea mantener (0 para finalizar): ")
        if seleccion.strip() == '0':
            break
        try:
            indices = [int(idx.strip()) - 1 for idx in seleccion.split(',')]
            for idx in indices:
                if 0 <= idx < len(encabezados) and encabezados[idx] not in columnas_a_mantener:
                    columnas_a_mantener.append(encabezados[idx])
                    print(f"Columna '{encabezados[idx]}' agregada.")
        except ValueError:
            print("Entrada inválida. Por favor, ingrese números válidos.")

    if not columnas_a_mantener:
        print("No se seleccionaron columnas. Finalizando.")
        return

    # Filtrar columnas seleccionadas
    try:
        df = filtrar_datos_para_mantener(df, columnas_a_mantener)
    except KeyError as e:
        print(f"Error al filtrar columnas: {e}")
        return

    # Validar números telefónicos en las columnas pertinentes
    df = validar_telefonos(df)

    # Reorganizar las columnas según el template
    df = reorganizar_columnas_template(df)

    # Mostrar una vista previa de los datos limpios
    print("\nDatos después de la limpieza:")
    print(df.head())

    # Guardar el archivo limpio
    while True:
        ruta_guardado = input("\nIngresa la ruta para guardar el archivo limpio (ejemplo: 'archivo_limpio.xlsx'): ")
        if ruta_guardado.endswith(".xlsx"):
            try:
                df.to_excel(ruta_guardado, index=False)
                print(f"Archivo guardado exitosamente en {ruta_guardado}")
                break
            except Exception as e:
                print(f"Error al guardar el archivo: {e}")
        else:
            print("Por favor, proporciona un nombre de archivo con la extensión .xlsx")

def reorganizar_columnas_template(df):
    """
    Reorganiza las columnas del DataFrame para que sigan el template:
    1. Nombre Completo
    2. Teléfono
    3. Email

    Si alguna columna está ausente, se agrega con valores predeterminados ("No hay <columna>").
    Además, renombra las columnas detectadas para ajustarse al template.
    """
    # Detectar columnas existentes
    nombre_col = [col for col in df.columns if "name" in col]
    telefono_col = [col for col in df.columns if df[col].astype(str).str.match(r"^\d{10}$").any()]
    email_col = [col for col in df.columns if df[col].astype(str).str.contains(r"@", na=False).any()]

    # Configurar nombres de las columnas en el template
    nombre_columna = "name"
    telefono_columna = "phone"
    email_columna = "email"

    # Renombrar las columnas detectadas
    if nombre_col:
        df.rename(columns={nombre_col[0]: nombre_columna}, inplace=True)
    else:
        df[nombre_columna] = "No hay Nombre"

    if telefono_col:
        df.rename(columns={telefono_col[0]: telefono_columna}, inplace=True)
    else:
        df[telefono_columna] = "No hay Teléfono"

    if email_col:
        df.rename(columns={email_col[0]: email_columna}, inplace=True)
    else:
        df[email_columna] = "No hay Email"

    # Crear el nuevo orden de columnas
    columnas_ordenadas = [nombre_columna, telefono_columna, email_columna]

    # Agregar las demás columnas
    otras_columnas = [col for col in df.columns if col not in columnas_ordenadas]
    columnas_ordenadas.extend(otras_columnas)

    # Reorganizar el DataFrame
    return df[columnas_ordenadas]

# Llamar a la función principal
if __name__ == "__main__":
    ruta = input("Ingresa la ruta del archivo Excel que deseas limpiar (ejemplo: 'archivo.xlsx'): ")
    limpiar_archivo_excel(ruta)