import aspose.words as aw
import os
import json
import re

# path
path = "G:\\My Drive\\MOTHIUS\\Documents\\Thrs\\Others\\Papeles\\Papeles - ELVER ANDRES TAO FLOREZ\\CSJ\\OnsiteFormatOne"
# get files into path

files = os.listdir(path)

for index, file in enumerate(files):
   #JSON
   dataFromDocx = {}
   
   
   docToRead = aw.Document(f"{path}\\{str(file)}")
   # Read all the contents from the node types paragraph
   contentText = []
   for paragraph in docToRead.get_child_nodes(aw.NodeType.PARAGRAPH, True):
      paragraph = paragraph.as_paragraph()
      # contentText.append((f"Paragraph {i} {paragraph.to_string(aw.SaveFormat.TEXT)}").strip())
      contentText.append((paragraph.to_string(aw.SaveFormat.TEXT)))

   # get and show custumerName (Nombre de Contacto)
   nombreDeContactoIndex = [i for i, s in enumerate(contentText) if "Nombre de Contacto" in s]
   custumerName = contentText[nombreDeContactoIndex[0]+1]

   # get and show number and case type (No. Caso Diagnóstico)
   casoDiagnosticoIndex = [i for i, s in enumerate(contentText) if "No. Caso" in s]
   case = re.search(r'\d+ (RQ|INC)',contentText[casoDiagnosticoIndex[0]]).group(0)


   # End Falla reportada
   # startFallaReportadaIndex = contentText.index("Falla Reportada")
   startFallaReportadaIndex = [i for i, s in enumerate(contentText) if "Falla Reportada" in s]
   # print(startFallaReportadaIndex)
   # print(f"{index+1}: {contentText[startFallaReportadaIndex]}")
   endFallaReportadaIndex = [i for i, s in enumerate(contentText) if "Ingeniero Asignado" in s]
   # print(custumerName) # Custumer Name
   # Save file content into variable contentText
   # contentText = "\n".join(contentText)
   contentText = "\n".join(contentText)
   builder0 = aw.DocumentBuilder()
   builder0.write(contentText)


   builder = aw.DocumentBuilder()
   builder.write(contentText)
   contentTextBuild = builder.document
   contentTextBuild.save(f".\\files\\{str(file)}.txt")
   # print(contentText)
   
   
   # -----------Test
   
   newDoc0 = builder0.document
   # print(newDoc0.get_text())

   # -----------End Test
   
   
   # read element
   with open(f".\\files\\{str(file)}.txt", encoding='utf-8') as file0:
      contenido = file0.read()
   # hacae la busqueda
   case = re.search('No. Caso Diagnóstico\s*(.*?)\n', contentText).group(1).strip()
   
   dataFromDocx = {
      "custumerName": case,
      "case": custumerName
   }
   print(dataFromDocx)
   
   


