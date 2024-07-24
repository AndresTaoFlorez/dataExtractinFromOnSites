import aspose.words as aw
import os
import json
import re

# path
# path = "G:\\My Drive\\MOTHIUS\\Documents\\Thrs\\Others\\Papeles\\Papeles - ELVER ANDRES TAO FLOREZ\\CSJ\\OnsiteFormatOne"
path = "C:\\Users\\Soporte\\Documents\\localweb\\AmyEaseKindBE\\public\\uploads"
# get files into path

files = os.listdir(path)

for index, file in enumerate(files):
   #JSON
   dataFromDocx = {}
   
   
   docToRead = aw.Document(f"{path}\\{str(file)}")
   # Read all the contents from the node types paragraph and save in a LIST
   contentText = []
   for paragraph in docToRead.get_child_nodes(aw.NodeType.PARAGRAPH, True):
      paragraph = paragraph.as_paragraph()
      contentText.append((paragraph.to_string(aw.SaveFormat.TEXT)).strip())

   # -------------------------------------------------------------------

   # get and show custumerName (Nombre de Contacto)
   nombreDeContactoIndex = [i for i, s in enumerate(contentText) if "Nombre de Contacto" in s]
   custumerName = contentText[nombreDeContactoIndex[0]+1]
   
   # get and show mail address (Correo Electrónico)
   correoElectronicoIndex = [i for i, s in enumerate(contentText) if "Correo Electrónico" in s]
   cusstumerEmail = contentText[correoElectronicoIndex[0]+1]

   # get and show number and case type (No. Caso Diagnóstico)
   casoDiagnosticoIndex = [i for i, s in enumerate(contentText) if "No. Caso" in s]
   case = re.search(r'\d+ (RQ|INC)',contentText[casoDiagnosticoIndex[0]]).group(0)

   # get and show number and case type (No. Caso Diagnóstico)
   # get and show dateOfRequest (Fecha de Solicitud)
   dateOfRequestIndex = [i for i, s in enumerate(contentText) if "Fecha de Solicitud" in s]
   dateOfRequest = contentText[dateOfRequestIndex[0]+1]
   
   # get and show timeOfRequest (Hora de Solicitud)
   timeOfRequestIndex = [i for i, s in enumerate(contentText) if "Hora de Solicitud" in s]
   timeOfRequest = contentText[timeOfRequestIndex[0]+1]

   # get and show custumerIdentification (Número de Cédula)
   custumerIdentificationIndex = [i for i, s in enumerate(contentText) if "Número de Cédula" in s]
   custumerIdentification = contentText[custumerIdentificationIndex[0]+1]

   # get and show floorName (Oficina o Juzgado)
   floorNameIndex = [i for i, s in enumerate(contentText) if "Oficina o Juzgado" in s]
   floorName = contentText[floorNameIndex[0]+1]

   # get and show caseIssue (Falla Reportada)
   #  start FallaReportadaIndex
   startFallaReportadaIndex = [i for i, s in enumerate(contentText) if "Falla Reportada" in s]
   #  End Falla reportada
   endFallaReportadaIndex = [i for i, s in enumerate(contentText) if "Ingeniero Asignado" in s]
   restTemp = endFallaReportadaIndex[0] - startFallaReportadaIndex[0] - 1
   restTemp0 = contentText[startFallaReportadaIndex[0]+1 : startFallaReportadaIndex[0]+1+restTemp]
   caseIssue = " ".join(restTemp0)
   
   # get and show dateOfAttention (Fecha de Atención)
   dateOfAttentionIndex = [i for i, s in enumerate(contentText) if "Fecha de Atención" in s]
   dateOfAttention = contentText[dateOfAttentionIndex[0]+1]
   
   # get and show timeOfAttention (Hora de Atención)
   timeOfAttentionIndex = [i for i, s in enumerate(contentText) if "Hora de Atención" in s]
   timeOfAttention = contentText[timeOfAttentionIndex[0]+1]
   
   # get and show engineerName (Ingeniero Asignado - Nombre)
   engineerNameIndex = timeOfAttentionIndex[0]+3
   engineerName = contentText[engineerNameIndex]
   
   # get and show placa (Placa Equipo)
   placaIndex = [i for i, s in enumerate(contentText) if "Placa Equipo" in s]
   placa = contentText[placaIndex[0]+1]
   
   # get and show serialNumber (Serial Equipo)
   serialNumberIndex = [i for i, s in enumerate(contentText) if "Serial Equipo" in s]
   serialNumber = contentText[serialNumberIndex[0]+1]

   # get and show manufacturer (Marca Equipo)
   manufacturerIndex = [i for i, s in enumerate(contentText) if "Marca equipo" in s]
   manufacturer = contentText[manufacturerIndex[0]+1]
   
   # get and show manufacturerModel (Modelo Equipo)
   manufacturerModelIndex = [i for i, s in enumerate(contentText) if "Modelo Equipo" in s]
   manufacturerModel = contentText[manufacturerModelIndex[0]+1]

   # get and show operatingSystem (Sistema Operativo)
   operatingSystemIndex = [i for i, s in enumerate(contentText) if "Sistema Operativo" in s]
   operatingSystem = contentText[operatingSystemIndex[0]+1]
   
   # get and show diagnostic (Información de Diagnóstico)
   #  start diagnostic
   startDiagnostic = [i for i, s in enumerate(contentText) if "Información de Diagnóstico:" in s]
   #  End Falla reportada
   endDiagnostic = [i for i, s in enumerate(contentText) if "Solución Entregada:" in s]
   restTemp = (endDiagnostic[0] - startDiagnostic[0]) - 1
   restTemp0 = contentText[startDiagnostic[0] : startDiagnostic[0]+1+restTemp]
   diagnostic = " ".join(restTemp0).replace('Información de Diagnóstico:', "").strip()

   # get and show solution (Solución Entregada)
   #  start solution
   startSolution = [i for i, s in enumerate(contentText) if "Solución Entregada:" in s]
   #  End Falla reportada
   endSolution = [i for i, s in enumerate(contentText) if "Observación por el Cliente" in s]
   restTemp = (endSolution[0] - startSolution[0]) - 1
   restTemp0 = contentText[startSolution[0] : startSolution[0]+1+restTemp]
   solution = " ".join(restTemp0).replace('Solución Entregada:', "").strip()

   # ------------------------------------------------------------------
   
   
   dataFromDocx = {
      "case": case,
      "custumerName": custumerName,
      "custumerIdentification": custumerIdentification,
      "custumerEmail": cusstumerEmail,
      "dateOfRequest": dateOfRequest,
      "timeOfRequest": timeOfRequest,
      "floorName": floorName,
      "caseIssue": caseIssue,
      "dateOfAttention": dateOfAttention,
      "timeOfAttention": timeOfAttention,
      "engineerName": engineerName,
      "placa": placa,
      "serialNumber": serialNumber,
      "manufacturer": manufacturer,
      "manufacturerModel": manufacturerModel,
      "operatingSystem": operatingSystem,
      "diagnostic": diagnostic,
      "solution": solution
   }
   
   with open(f".\\files\\{file.replace('.docx','')}.json", "w", encoding='utf-8') as file0:
      json.dump(dataFromDocx, file0, ensure_ascii=False)
   print(dataFromDocx)
   
   


