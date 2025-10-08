function descriptionnombreggp() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("==No hay filas válidas para procesar.");
    return;
  }
  for (const rowNumber of selectedRows) {
    console.log("------------------------------", rowNumber);
    nombregp(rowNumber, sheet);
  }
}

function clasificarPreguntasPorCompetencia() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columna para clasificación si no existe
  sheet.createColumn(["Clasificación Competencia", "Justificación Clasificación"]);
  
  for (const rowNumber of selectedRows) {
    console.log("Clasificando pregunta en fila:", rowNumber);
    clasificarPregunta(rowNumber, sheet, chatGPTClient);
  }
}

function nombregp(rowNumber, sheet) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    Clase: sheet.getCellValue("Clase de Ítem", rowNumber),
    Clave_Respuesta: sheet.getCellValue("Clave", rowNumber)
  };
  const chatGPTClient = new ChatGPT();
  prompt_requirement = `
Te voy a dar el contexto de un proyecto, sobre el cual estoy trabajando:

Detalle de rubro:
${data.Detalle}

# TAREA
quiero que me ayudes a generar el nombre del rubro basado en la información del detalle  del item del presupuesto que puede ser personal o insumos que se va a contratar o a comprar, escribelo de forma humanizado en maximo 12 palabras minimo 3 palabras: 

Nunca inventes datos ni agregues nombres de equipos o no se mencionen en la información.

Nunca colocar nombres de personas.

Escribir de forma tecnica y formal.

No utilizar los dos puntos, si no que escribir de forma continua.

escribir de forma explicita y no general.

Ahora te coloco el texto para modificar:

${data.Nombre}

Devuelve el texto modificado con la justificación que cumpla todas las reglas.
Responde solo con el siguiente JSON:

{
  "Nombre": "......",
}
Redacta ahora:
`;
  let prompt = prompt_requirement;
  let result = chatGPTClient.request(prompt, data["item"], null, null);
  console.log(result);
  sheet.saveResultToSheetRow(result, rowNumber, "Nombre");
}
