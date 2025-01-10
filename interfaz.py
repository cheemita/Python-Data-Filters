import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import re

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

def validar_telefonos(df):
    for columna in df.columns:
        muestra = df[columna].astype(str).head(10).tolist()

        def es_posible_telefono(valor):
            valor = re.sub(r'\D', '', valor)
            return len(valor) == 10 and valor[:3] in CODIGOS_AREA_VALIDOS

        es_columna_telefonica = sum([es_posible_telefono(valor) for valor in muestra]) > (len(muestra) // 2)

        if es_columna_telefonica:
            df[columna] = df[columna].astype(str).apply(lambda x: re.sub(r'\D', '', x) if len(re.sub(r'\D', '', x)) == 10 else None)
            df = df[df[columna].notna()]

    return df

def concatenar_nombres_apellidos(df, cols):
    df["Nombre Completo"] = df[cols].astype(str).agg(' '.join, axis=1)
    df = df.drop(columns=cols)
    return df

def procesar_nombre_completo(df, col_nombre_completo):
    try:
        df["Nombre Completo"] = df[col_nombre_completo].astype(str)
        df.drop(columns=[col_nombre_completo], inplace=True)
        return df
    except Exception as e:
        print(f"Error al procesar el nombre completo: {e}")
        return df

def detectar_columnas_nombres(df):
    # Detectamos columnas posibles para nombres y apellidos
    posibles_columnas = []
    for col in df.columns:
        muestra = df[col].astype(str).head(10).tolist()
        if all(re.match(r'^[A-Za-záéíóúÁÉÍÓÚñÑ ]+$', str(valor)) and len(str(valor).split()) == 1 for valor in muestra if pd.notna(valor)):
            posibles_columnas.append(col)
    return posibles_columnas


class ExcelCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Cleaner")

        self.df = None
        self.file_path = ""

        self.setup_ui()

    def setup_ui(self):
        self.btn_select_file = tk.Button(self.root, text="Seleccionar Archivo", command=self.select_file)
        self.btn_select_file.pack(pady=10)

        self.columns_frame = tk.LabelFrame(self.root, text="Seleccione Columnas a Mantener")
        self.columns_frame.pack(pady=10, fill="both", expand=True)

        self.columns_list = tk.Listbox(self.columns_frame, selectmode="multiple", width=50, height=10)
        self.columns_list.pack(side="left", fill="both", expand=True)

        self.scrollbar = tk.Scrollbar(self.columns_frame, orient="vertical", command=self.columns_list.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.columns_list.config(yscrollcommand=self.scrollbar.set)

        self.btn_save_file = tk.Button(self.root, text="Guardar Archivo", command=self.save_file)
        self.btn_save_file.pack(pady=5)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not self.file_path:
            return

        try:
            self.df = pd.read_excel(self.file_path, header=None)
            first_row = self.df.iloc[0].dropna().tolist()

            if all(isinstance(item, str) for item in first_row):
                self.df.columns = self.df.iloc[0]
                self.df = self.df[1:]
            else:
                self.df.columns = [f"Columna_{i+1}" for i in range(self.df.shape[1])]

            self.process_data()
            self.update_columns_list()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

    def preguntar_concatenar_nombres(self, posibles_columnas):
        # Preguntamos si el usuario desea concatenar las columnas seleccionadas como 'Nombre Completo'
        respuesta = messagebox.askyesno("Concatenar Nombres y Apellidos", "¿Deseas concatenar las siguientes columnas como 'Nombre Completo'?")
        
        if respuesta:
            # Permitimos seleccionar las columnas que se desean concatenar
            selected_cols = simpledialog.askstring("Seleccionar columnas", f"Selecciona las columnas a concatenar de las siguientes: {', '.join(posibles_columnas)}")
            
            if selected_cols:
                # Convertimos la entrada en una lista de columnas seleccionadas
                selected_cols = [col.strip() for col in selected_cols.split(',')]
                
                # Verificamos que todas las columnas seleccionadas estén en el dataframe
                valid_cols = [col for col in selected_cols if col in posibles_columnas]
                if len(valid_cols) == len(selected_cols):  # Aseguramos que todas las columnas sean válidas
                    return valid_cols
                else:
                    messagebox.showerror("Error", "Algunas de las columnas seleccionadas no son válidas.")
                    return None
            else:
                messagebox.showwarning("Advertencia", "No se ha ingresado ninguna columna. La operación se cancelará.")
                return None
        
        return None
    
    def process_data(self):
        try:
            # Detectamos las columnas posibles para concatenar
            posibles_columnas = detectar_columnas_nombres(self.df)

            # Preguntamos si el usuario desea concatenar y, si es así, qué columnas
            selected_cols = self.preguntar_concatenar_nombres(posibles_columnas)
            
            if selected_cols:
                # Concatenamos las columnas seleccionadas
                self.df = concatenar_nombres_apellidos(self.df, selected_cols)

            # Si ya hay una columna "Nombre Completo", procesamos ese caso
            elif "Nombre Completo" in self.df.columns:
                col_nombre_completo = "Nombre Completo"
                self.df = procesar_nombre_completo(self.df, col_nombre_completo)

            # Aseguramos que "Nombre Completo" esté presente y sea la primera columna
            if "Nombre Completo" in self.df.columns:
                cols = ["Nombre Completo"] + [col for col in self.df.columns if col != "Nombre Completo"]
                self.df = self.df[cols]
            
            # Validamos los números de teléfono
            self.df = validar_telefonos(self.df)

        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar los datos: {e}")

    def update_columns_list(self):
        self.columns_list.delete(0, tk.END)
        for col in self.df.columns:
            self.columns_list.insert(tk.END, col)

    def save_file(self):
        if self.df is None:
            messagebox.showerror("Error", "No hay datos para guardar.")
            return

        selected_indices = self.columns_list.curselection()
        if not selected_indices:
            messagebox.showerror("Error", "Debe seleccionar al menos una columna para guardar.")
            return

        selected_columns = [self.columns_list.get(i) for i in selected_indices]
        self.df = self.df[selected_columns]

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return

        try:
            self.df.to_excel(save_path, index=False)
            messagebox.showinfo("Éxito", f"Archivo guardado en {save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()