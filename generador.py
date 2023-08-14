from pathlib import Path  # core python module
import win32com.client as win32
import pandas as pd
from datetime import datetime
from babel.dates import format_date
import re
import os
import shutil

# Path settings
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "input"
output_dir = current_dir / "certificados"

if os.path.exists(output_dir) and os.path.isdir(output_dir):
    # Eliminar el directorio
    shutil.rmtree(output_dir)

output_dir.mkdir(parents=True, exist_ok=True)
direcion=''
for doc_file in Path(input_dir).rglob("*.xlsx*"):
    direcion=str(doc_file)

df = pd.read_excel(direcion, sheet_name="ASISTENTES", skiprows=3, header=0)
ev=''
while True:
    ev=input("Ingrese nombre del evento:\n")
    if ev in df["EVENTO"].values:
        break
    print("Evento no encontrado")
df=df[df["EVENTO"]==ev]
umbral = df.shape[1] - 1 
df = df.dropna(thresh=umbral)
#df=df.head()
fecha=df["FECHA"].values
horario=df["HORARIO"].values
evento=df["EVENTO"].values
n=df["Nº"].values
participante=df["PARTICIPANTE"].values
codigo=df["CÓDIGO"].values
codigo_barra=df["CÓDIGO DE BARRAS"].values

# Find & replace
wd_replace = 2  # 2=replace all occurences, 1=replace one occurence, 0=replace no occurences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached

# Open Word
word_app = word_app = win32.gencache.EnsureDispatch("Word.Application")
word_app.Visible = False
word_app.DisplayAlerts = False
# Open each document and replace string
for doc_file in Path(input_dir).rglob("*.doc*"):
    direcion=str(doc_file)
for k in range(df.shape[0]):
    doc = word_app.Documents.Open(direcion)
    fecha_objeto = datetime.strptime(str(fecha[k]), "%Y-%m-%d %H:%M:%S")
    fecha_formateada = format_date(fecha_objeto, format="d 'de' MMMM 'del' yyyy", locale='es')
    reemplazos = {
        'arg1': participante[k],
        'arg2': evento[k],
        'arg3':fecha_formateada,
        'arg4':codigo[k],
        'Arg5':codigo_barra[k]
    }
    patron = '|'.join(rf'\b{re.escape(key)}\b' for key in reemplazos.keys())
    # VBA SO reference: https://stackoverflow.com/a/26266598
    # Loop through all the shapes
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:  # Verificar si el shape tiene un TextFrame (cuadro de texto)
            texto_del_cuadro = shape.TextFrame.TextRange.Text
            n=shape.TextFrame.TextRange.ParagraphFormat.Alignment
            nueva_cadena = re.sub(patron, lambda match: reemplazos[match.group()], texto_del_cuadro, flags=re.IGNORECASE)
            shape.TextFrame.TextRange.Text=nueva_cadena
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = n

    # Save the new file
    output_path = output_dir / f"{participante[k]}.pdf"
    doc.SaveAs(str(output_path), FileFormat=17)
    doc.Close(SaveChanges=False)
    print("%i%%, N°: %i"%(100*(k+1)/df.shape[0],k+1))
word_app.Quit()

