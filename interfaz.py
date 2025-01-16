import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re

# Lista de códigos de área válidos (solo EE. UU. y Canadá)
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

# Función para validar números telefónicos
def validar_telefonos(df):
    """
    Valida y limpia números telefónicos en las columnas pertinentes.
    Detecta automáticamente columnas que contienen números de teléfono,
    y elimina registros con números inválidos o filas donde no hay valores.
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
            df[columna] = df[columna].astype(str)

            def limpiar_y_validar(numero):
                """
                Limpia y valida un número telefónico.
                Retorna el número limpio si es válido, o None si no lo es.
                """
                numero = re.sub(r'\D', '', numero)  # Eliminar caracteres no numéricos
                if len(numero) != 10 or numero[:3] not in CODIGOS_AREA_VALIDOS:  # Validar formato y código de área
                    return None
                return numero  # Retornar el número limpio si es válido

            # Limpiar y validar la columna
            df[f"{columna}_limpio"] = df[columna].apply(limpiar_y_validar)

            # Mostrar estadísticas
            total_validos = df[f"{columna}_limpio"].notna().sum()
            total_invalidos = len(df) - total_validos
            print(f" - Total válidos: {total_validos}")
            print(f" - Total inválidos: {total_invalidos}")

            # Filtrar registros válidos
            df = df[df[f"{columna}_limpio"].notna()]

            # Reemplazar la columna original por los números limpios
            df[columna] = df[f"{columna}_limpio"]
            df = df.drop(columns=[f"{columna}_limpio"])

    return df

# Función para concatenar nombres y apellidos
def concatenar_nombres_apellidos(df, col_nombre, col_apellido):
    df[col_nombre] = df[col_nombre].fillna("")  # Rellenar valores nulos
    df[col_apellido] = df[col_apellido].fillna("")
    df["name"] = df[col_nombre].astype(str) + " " + df[col_apellido].astype(str)
    df = df[~((df[col_nombre] == "") & (df[col_apellido] == ""))]  # Eliminar filas donde ambos estén vacíos
    return df

# Función para procesar nombre completo en una sola columna
def procesar_nombre_completo(df, col_nombre_completo):
    try:
        df["name"] = df[col_nombre_completo].astype(str)  # Usar la columna completa como 'name'
        df.drop(columns=[col_nombre_completo], inplace=True)  # Eliminar la columna de nombre completo
        return df
    except Exception as e:
        print(f"Error al procesar el nombre completo: {e}")
        return df


# Función para filtrar las columnas seleccionadas por el usuario
def filtrar_datos_para_mantener(df, columnas_a_mantener):
    return df[columnas_a_mantener]  # Solo mantener las columnas seleccionadas

# Función para reorganizar las columnas según un formato predeterminado
def reorganizar_columnas_template(df):
    nombre_columna = "name"
    telefono_columna = "phone"
    email_columna = "email"

    # Buscar las columnas correspondientes a nombre, teléfono y email
    nombre_col = [col for col in df.columns if "name" in col]
    telefono_col = [col for col in df.columns if df[col].astype(str).str.match(r"^\d{10}$").any()]
    email_col = [col for col in df.columns if df[col].astype(str).str.contains(r"@", na=False).any()]

    # Renombrar las columnas si se encuentran
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

    # Reorganizar las columnas según el orden deseado
    columnas_ordenadas = [nombre_columna, telefono_columna, email_columna]
    otras_columnas = [col for col in df.columns if col not in columnas_ordenadas]
    columnas_ordenadas.extend(otras_columnas)

    return df[columnas_ordenadas]

# Función principal para cargar y procesar el archivo Excel
def procesar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccione un archivo Excel", filetypes=[("Archivos Excel", "*.xlsx")])
    if not archivo:
        return

    try:
        encabezados, df = leer_archivo_excel(archivo)
        print("Archivo cargado con éxito.")

        def procesar_nombres():
            opcion = var_opcion.get()
            if opcion == 1:  # Concatenar nombres y apellidos
                col_nombre = combo_nombre.get()
                col_apellido = combo_apellido.get()
                df_resultado = concatenar_nombres_apellidos(df, col_nombre, col_apellido)
            elif opcion == 2:  # Usar columna de nombre completo
                col_nombre_completo = combo_nombre_completo.get()
                df_resultado = procesar_nombre_completo(df, col_nombre_completo)
            else:
                messagebox.showerror("Error", "Seleccione una opción válida.")
                return

            # Validar números telefónicos si el usuario lo solicita
            if var_validar_telefonos.get():
                df_resultado = validar_telefonos(df_resultado)

            columnas_seleccionadas = [col for col in lista_columnas.curselection()]
            columnas_a_mantener = [encabezados[idx] for idx in columnas_seleccionadas]
            columnas_a_mantener.append("name")

            df_resultado = filtrar_datos_para_mantener(df_resultado, columnas_a_mantener)
            df_resultado = reorganizar_columnas_template(df_resultado)

            ruta_guardado = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel", "*.xlsx")])
            if ruta_guardado:
                df_resultado.to_excel(ruta_guardado, index=False)
                messagebox.showinfo("Éxito", f"Archivo guardado en {ruta_guardado}")

        # Crear ventana emergente
        ventana = tk.Toplevel()
        ventana.title("Procesar Archivo Excel")

        tk.Label(ventana, text="Seleccione cómo procesar los nombres:").pack()
        var_opcion = tk.IntVar()

        # Opción para concatenar nombres y apellidos
        radio_concatenar = tk.Radiobutton(ventana, text="Concatenar columnas de Nombre y Apellido", variable=var_opcion, value=1)
        radio_concatenar.pack()

        # Opción para usar una columna de nombre completo
        radio_completo = tk.Radiobutton(ventana, text="Usar una columna para el Nombre Completo", variable=var_opcion, value=2)
        radio_completo.pack()

        # Variables para las columnas a seleccionar
        combo_nombre = tk.StringVar(value=encabezados)
        combo_apellido = tk.StringVar(value=encabezados)
        combo_nombre_completo = tk.StringVar(value=encabezados)

        # Crear paneles para columnas de nombre y apellido, inicialmente ocultos
        frame_columnas = tk.Frame(ventana)
        tk.Label(frame_columnas, text="Seleccione las columnas para concatenar:").pack()
        combo_nombre_menu = tk.OptionMenu(frame_columnas, combo_nombre, *encabezados)
        combo_nombre_menu.pack()
        combo_apellido_menu = tk.OptionMenu(frame_columnas, combo_apellido, *encabezados)
        combo_apellido_menu.pack()

        frame_nombre_completo = tk.Frame(ventana)
        tk.Label(frame_nombre_completo, text="Seleccione la columna de Nombre Completo:").pack()
        combo_nombre_completo_menu = tk.OptionMenu(frame_nombre_completo, combo_nombre_completo, *encabezados)
        combo_nombre_completo_menu.pack()

        def actualizar_opciones():
            if var_opcion.get() == 1:  # Concatenar columnas
                frame_columnas.pack()  # Mostrar columnas de nombre y apellido
                frame_nombre_completo.pack_forget()  # Ocultar nombre completo
            elif var_opcion.get() == 2:  # Usar nombre completo
                frame_columnas.pack_forget()  # Ocultar columnas de nombre y apellido
                frame_nombre_completo.pack()  # Mostrar nombre completo

        # Llamar a la función de actualización cada vez que se cambie la selección
        radio_concatenar.config(command=actualizar_opciones)
        radio_completo.config(command=actualizar_opciones)

        # Llamar a la función inicial para actualizar las opciones
        actualizar_opciones()

        # Selección de columnas a mantener
        tk.Label(ventana, text="Seleccione las columnas a mantener:").pack()
        lista_columnas = tk.Listbox(ventana, selectmode=tk.MULTIPLE)
        for col in encabezados:
            lista_columnas.insert(tk.END, col)
        lista_columnas.pack()

        # Checkbox para validar números telefónicos
        var_validar_telefonos = tk.IntVar()
        tk.Checkbutton(ventana, text="Validar números telefónicos", variable=var_validar_telefonos).pack()

        # Botón para procesar el archivo
        tk.Button(ventana, text="Procesar y Guardar", command=procesar_nombres).pack()

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar el archivo: {e}")

# Interfaz principal
root = tk.Tk()
root.title("Limpieza Avanzada de Archivos Excel")

btn_cargar = tk.Button(root, text="Cargar Archivo Excel", command=procesar_archivo)
btn_cargar.pack(pady=20)

root.mainloop()
