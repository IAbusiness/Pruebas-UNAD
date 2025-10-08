function evaluarPreguntasCalidad() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columnas para la evaluación si no existen
  sheet.createColumn([
    "Evaluación Temario", 
    "Evaluación Nivel Universitario", 
    "Evaluación Concordancia Clase",
    "Puntaje Total Evaluación",
    "Observaciones Evaluación"
  ]);
  
  for (const rowNumber of selectedRows) {
    console.log("Evaluando pregunta en fila:", rowNumber);
    evaluarPregunta(rowNumber, sheet, chatGPTClient);
  }
}

function evaluarPregunta(rowNumber, sheet, chatGPTClient) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    Clase: sheet.getCellValue("Clase de Ítem", rowNumber) || sheet.getCellValue("Clase", rowNumber),
    Pregunta: sheet.getCellValue("Enunciado (con el contexto)", rowNumber),
    OpcionA: sheet.getCellValue("Texto A (con el contexto)", rowNumber) || "",
    OpcionB: sheet.getCellValue("Texto B (con el contexto)", rowNumber) || "",
    OpcionC: sheet.getCellValue("Texto C (con el contexto)", rowNumber) || "",
    OpcionD: sheet.getCellValue("Texto D (con el contexto)", rowNumber) || "",
    ClaveRespuesta: sheet.getCellValue("Clave", rowNumber) || sheet.getCellValue("Clave de Respuesta", rowNumber),
    Competencia: sheet.getCellValue("Competencia", rowNumber) || ""
  };

  // Definir criterios de evaluación para nivel universitario
  const criteriosUniversitarios = {
    complejidad: "Requiere pensamiento crítico, análisis profundo y aplicación de conceptos teóricos avanzados",
    vocabulario: "Utiliza terminología técnica y académica apropiada para el nivel superior",
    abstraccion: "Involucra conceptos abstractos, teorías complejas y relaciones multivariables",
    aplicacion: "Conecta teoría con práctica profesional y situaciones del mundo real",
    rigor: "Mantiene precisión académica y rigurosidad científica en planteamientos"
  };

  const prompt_evaluacion = `
Eres un experto evaluador académico especializado en educación superior. Tu tarea es evaluar de manera integral una pregunta de examen universitario en tres dimensiones críticas.

CONTEXTO EDUCATIVO:
- Materia: ${data.Materia}
- Temario: ${data.Temario}
- Competencia: ${data.Competencia}
- Clase de Competencia Declarada: ${data.Clase}

PREGUNTA A EVALUAR:
${data.Pregunta}

OPCIONES:
A) ${data.OpcionA}
B) ${data.OpcionB}
C) ${data.OpcionC}
D) ${data.OpcionD}

Respuesta Correcta: ${data.ClaveRespuesta}

CRITERIOS DE EVALUACIÓN:

## 1. CONCORDANCIA CON EL TEMARIO (Puntaje: 0-100)
Evalúa si la pregunta:
- Aborda específicamente contenidos del temario "${data.Temario}"
- Es relevante para la materia "${data.Materia}"
- Se alinea con la competencia "${data.Competencia}"
- No incluye contenidos ajenos o fuera del alcance temático

## 2. NIVEL UNIVERSITARIO (Puntaje: 0-100)
Evalúa si la pregunta cumple con estándares universitarios:
- **Complejidad**: ${criteriosUniversitarios.complejidad}
- **Vocabulario**: ${criteriosUniversitarios.vocabulario}
- **Abstracción**: ${criteriosUniversitarios.abstraccion}
- **Aplicación**: ${criteriosUniversitarios.aplicacion}
- **Rigor**: ${criteriosUniversitarios.rigor}

## 3. CONCORDANCIA CON LA CLASE DE COMPETENCIA (Puntaje: 0-100)
Evalúa si la pregunta efectivamente requiere el tipo de proceso mental de la clase "${data.Clase}":

### Características esperadas según la clase:
- **INTERPRETATIVA**: Comprensión, traducción de significados, descripción de variables
- **ARGUMENTATIVA**: Justificación, análisis teórico, evaluación de planteamientos
- **PROPOSITIVA**: Resolución de problemas, planteamiento de hipótesis
- **RECONOCIMIENTO**: Memoria de datos, fechas, definiciones
- **COMPRENSIÓN**: Organización, clasificación, relación de información
- **APLICACIÓN**: Uso de reglas, métodos, teorías en situaciones específicas
- **ANÁLISIS**: Descomposición, identificación de partes y relaciones
- **EVALUACIÓN**: Juicios de valor, decisiones fundamentadas
- **SÍNTESIS**: Integración, creación, combinación de elementos

INSTRUCCIONES DE EVALUACIÓN:
- Asigna puntajes de 0-100 para cada criterio
- Sé objetivo y fundamenta cada puntaje
- Identifica fortalezas y debilidades específicas
- Sugiere mejoras concretas cuando sea necesario

Responde ÚNICAMENTE con el siguiente formato JSON:

{
  "Evaluación Temario": 85,
  "Evaluación Nivel Universitario": 78,
  "Evaluación Concordancia Clase": 92,
  "Puntaje Total Evaluación": 85,
  "Observaciones Evaluación": "FORTALEZAS: [Lista específica de aspectos positivos] | DEBILIDADES: [Lista específica de aspectos a mejorar] | RECOMENDACIONES: [Sugerencias concretas para optimizar la pregunta]"
}

IMPORTANTE: 
- El Puntaje Total debe ser el promedio de los tres criterios
- Las observaciones deben ser específicas y constructivas
- Incluye ejemplos concretos cuando sea relevante
`;

  try {
    let result = chatGPTClient.request(prompt_evaluacion, "", null, null);
    console.log("Resultado evaluación:", result);
    
    // Guardar los resultados en las columnas correspondientes
    sheet.saveResultToSheetRow(result, rowNumber, "Evaluación Temario");
    sheet.saveResultToSheetRow(result, rowNumber, "Evaluación Nivel Universitario");
    sheet.saveResultToSheetRow(result, rowNumber, "Evaluación Concordancia Clase");
    sheet.saveResultToSheetRow(result, rowNumber, "Puntaje Total Evaluación");
    sheet.saveResultToSheetRow(result, rowNumber, "Observaciones Evaluación");
    
  } catch (error) {
    console.error("Error evaluando pregunta:", error);
    SpreadsheetApp.getUi().alert("Error al evaluar la pregunta: " + error.message);
  }
}
