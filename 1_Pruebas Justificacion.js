function generarRealimentacionRespuestas() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columnas para la realimentación si no existen
  sheet.createColumn([
    "Realimentación A",
    "Realimentación B", 
    "Realimentación C",
    "Realimentación D"
  ]);
  
  let preguntasProcesadas = 0;
  
  for (const rowNumber of selectedRows) {
    console.log(`Generando realimentación para pregunta en fila: ${rowNumber}`);
    generarJustificacionOpciones(rowNumber, sheet, chatGPTClient);
    preguntasProcesadas++;
  }
  
  // Mostrar resumen del procesamiento
  const mensaje = `Realimentación generada para ${preguntasProcesadas} preguntas`;
  SpreadsheetApp.getUi().alert(mensaje);
  console.log(mensaje);
}

function generarJustificacionOpciones(rowNumber, sheet, chatGPTClient) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    Clase: sheet.getCellValue("Clase de Ítem", rowNumber) || sheet.getCellValue("Clase", rowNumber),
    Competencia: sheet.getCellValue("Competencia", rowNumber) || "",
    
    // Usar las preguntas mejoradas si existen, si no usar las originales
    Pregunta: sheet.getCellValue("Enunciado (con el contexto)", rowNumber),
    OpcionA: sheet.getCellValue("Texto A (con el contexto)", rowNumber) || "",
    OpcionB: sheet.getCellValue("Texto B (con el contexto)", rowNumber) || "",
    OpcionC: sheet.getCellValue("Texto C (con el contexto)", rowNumber) || "",
    OpcionD: sheet.getCellValue("Texto D (con el contexto)", rowNumber) || "",
    ClaveRespuesta: sheet.getCellValue("Respuesta Correcta", rowNumber) || 
                    sheet.getCellValue("Clave", rowNumber) || 
                    sheet.getCellValue("Clave de Respuesta", rowNumber)
  };

  const prompt_realimentacion = `
Eres un experto pedagogo universitario especializado en crear realimentación educativa de alta calidad. Tu tarea es generar justificaciones específicas para cada opción de respuesta de una pregunta de examen universitario.

CONTEXTO EDUCATIVO:
- Materia: ${data.Materia}
- Temario: ${data.Temario}
- Competencia: ${data.Competencia}
- Clase de Competencia: ${data.Clase}

PREGUNTA:
${data.Pregunta}

OPCIONES DE RESPUESTA:
A) ${data.OpcionA}
B) ${data.OpcionB}
C) ${data.OpcionC}
D) ${data.OpcionD}

RESPUESTA CORRECTA: ${data.ClaveRespuesta}

## INSTRUCCIONES PARA LA REALIMENTACIÓN:

### PARA LA RESPUESTA CORRECTA (${data.ClaveRespuesta}):
- **EXALTAR POR QUÉ ES CORRECTA**: Explica claramente los fundamentos teóricos, conceptos o principios que hacen que esta opción sea la correcta
- **REFORZAR EL APRENDIZAJE**: Conecta la respuesta con el conocimiento específico del temario
- **DESTACAR ASPECTOS CLAVE**: Resalta los elementos más importantes que el estudiante debe recordar
- **TONO POSITIVO**: Usa un lenguaje que reconozca y refuerce el conocimiento correcto

### PARA LAS RESPUESTAS INCORRECTAS:
- **EXPLICAR POR QUÉ ESTÁ INCORRECTA**: Identifica específicamente el error conceptual o malentendido
- **CORREGIR MISCONCEPCIONES**: Aclara los conceptos que pueden estar confundiendo al estudiante
- **ORIENTAR HACIA LA RESPUESTA CORRECTA**: Guía sutilmente hacia el conocimiento correcto sin dar la respuesta directamente
- **TONO CONSTRUCTIVO**: Usa un enfoque educativo que ayude al aprendizaje sin desalentar

### CARACTERÍSTICAS DE UNA BUENA REALIMENTACIÓN:
1. **ESPECÍFICA**: Se refiere directamente al contenido de la opción
2. **EDUCATIVA**: Enseña o refuerza conceptos importantes
3. **FUNDAMENTADA**: Basada en principios teóricos del temario
4. **CLARA**: Lenguaje accesible pero técnicamente preciso
5. **CONSTRUCTIVA**: Ayuda al estudiante a mejorar su comprensión

### ESTRUCTURA RECOMENDADA:
- **Para correcta**: "✓ CORRECTO: [Explicación de por qué es correcta] + [Fundamento teórico] + [Conexión con el temario]"
- **Para incorrectas**: "✗ INCORRECTO: [Explicación del error] + [Aclaración conceptual] + [Orientación hacia el conocimiento correcto]"

### EJEMPLOS DE BUENA REALIMENTACIÓN:

**Para respuesta correcta:**
"✓ CORRECTO: Esta opción identifica correctamente el principio de [concepto] establecido en [teoría/autor]. Este fundamento es esencial en [materia] porque [explicación de importancia]. La aplicación de este concepto se observa cuando [ejemplo práctico]."

**Para respuesta incorrecta:**
"✗ INCORRECTO: Esta opción confunde [concepto A] con [concepto B]. Aunque ambos están relacionados con [tema], se diferencian en que [explicación de diferencia]. El concepto correcto para esta situación es [orientación sin dar respuesta directa]."

**Reglas:**
A) escribe en forma de parrafos continuos sin viñetas ni numeración.
B) escribe en maximo 40 palabras por justificación.

Responde ÚNICAMENTE con el siguiente formato JSON:

{
  "Realimentación A": "Justificación educativa específica para la opción A, siguiendo las instrucciones según sea correcta o incorrecta",
  "Realimentación B": "Justificación educativa específica para la opción B, siguiendo las instrucciones según sea correcta o incorrecta",
  "Realimentación C": "Justificación educativa específica para la opción C, siguiendo las instrucciones según sea correcta o incorrecta",
  "Realimentación D": "Justificación educativa específica para la opción D, siguiendo las instrucciones según sea correcta o incorrecta"
}
`;

  try {
    let result = chatGPTClient.request(prompt_realimentacion, "", null, null);
    console.log("Resultado realimentación:", result);
    
    // Guardar los resultados en las columnas correspondientes
    sheet.saveResultToSheetRow(result, rowNumber, "Realimentación A");
    sheet.saveResultToSheetRow(result, rowNumber, "Realimentación B");
    sheet.saveResultToSheetRow(result, rowNumber, "Realimentación C");
    sheet.saveResultToSheetRow(result, rowNumber, "Realimentación D");
    
  } catch (error) {
    console.error("Error generando realimentación:", error);
    SpreadsheetApp.getUi().alert("Error al generar realimentación: " + error.message);
  }
}
