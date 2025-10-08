function traducirPreguntasAIngles() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columnas para la traducción y etiquetas si no existen
  sheet.createColumn([
    "ENUNCIADO (EN)",
    "CLAVE (EN)", 
    "TEXTO A (EN)",
    "REALIMENTACIÓN A (EN)",
    "TEXTO B (EN)",
    "REALIMENTACIÓN B (EN)",
    "TEXTO C (EN)",
    "REALIMENTACIÓN C (EN)",
    "TEXTO D (EN)",
    "REALIMENTACIÓN D (EN)",
    "Etiqueta clase",
    "Etiqueta competencia"
  ]);
  
  let preguntasProcesadas = 0;
  
  for (const rowNumber of selectedRows) {
    console.log(`Traduciendo pregunta en fila: ${rowNumber}`);
    traducirPreguntaCompleta(rowNumber, sheet, chatGPTClient);
    aplicarEtiquetasNumericas(rowNumber, sheet);
    preguntasProcesadas++;
  }
  
  // Mostrar resumen del procesamiento
  const mensaje = `Traducción completada para ${preguntasProcesadas} preguntas`;
  SpreadsheetApp.getUi().alert(mensaje);
  console.log(mensaje);
}

function traducirPreguntaCompleta(rowNumber, sheet, chatGPTClient) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    
    // Datos a traducir
    Enunciado: sheet.getCellValue("Enunciado (con el contexto)", rowNumber) || "",
    Clave: sheet.getCellValue("Clave", rowNumber) || "",
    TextoA: sheet.getCellValue("Texto A (con el contexto)", rowNumber) || "",
    RealimentacionA: sheet.getCellValue("Realimentación A", rowNumber) || sheet.getCellValue("Realimentación A", rowNumber) || "",
    TextoB: sheet.getCellValue("Texto B (con el contexto)", rowNumber) || "",
    RealimentacionB: sheet.getCellValue("Realimentación B", rowNumber) || sheet.getCellValue("Realimentación B", rowNumber) || "",
    TextoC: sheet.getCellValue("Texto C (con el contexto)", rowNumber) || "",
    RealimentacionC: sheet.getCellValue("Realimentación C", rowNumber) || sheet.getCellValue("Realimentación C", rowNumber) || "",
    TextoD: sheet.getCellValue("Texto D (con el contexto)", rowNumber) || "",
    RealimentacionD: sheet.getCellValue("Realimentación D", rowNumber) || sheet.getCellValue("Realimentación D", rowNumber) || ""
  };

  const prompt_traduccion = `
Eres un traductor especializado en contenido académico universitario con expertise en traducción al inglés americano de Estados Unidos. Tu tarea es traducir con precisión académica una pregunta de examen completa con todas sus opciones y realimentaciones.

CONTEXTO ACADÉMICO:
- Materia: ${data.Materia}
- Temario: ${data.Temario}

CONTENIDO A TRADUCIR:

ENUNCIADO:
${data.Enunciado}

CLAVE DE RESPUESTA:
${data.Clave}

OPCIÓN A:
${data.TextoA}

REALIMENTACIÓN A:
${data.RealimentacionA}

OPCIÓN B:
${data.TextoB}

REALIMENTACIÓN B:
${data.RealimentacionB}

OPCIÓN C:
${data.TextoC}

REALIMENTACIÓN C:
${data.RealimentacionC}

OPCIÓN D:
${data.TextoD}

REALIMENTACIÓN D:
${data.RealimentacionD}

## INSTRUCCIONES DE TRADUCCIÓN:

### ESTILO Y REGISTRO:
- **INGLÉS AMERICANO**: Usar específicamente variantes, spelling y expresiones de Estados Unidos
- **REGISTRO ACADÉMICO**: Mantener formalidad y precisión técnica universitaria
- **TERMINOLOGÍA ESPECIALIZADA**: Usar términos técnicos apropiados para la disciplina académica
- **CLARIDAD**: Preservar la claridad y comprensibilidad del contenido original

### ASPECTOS TÉCNICOS:
- **EQUIVALENCIA CONCEPTUAL**: Mantener el mismo significado académico y técnico
- **COHERENCIA TERMINOLÓGICA**: Usar los mismos términos técnicos consistentemente
- **ADAPTACIÓN CULTURAL**: Ajustar referencias culturales cuando sea necesario
- **PRECISIÓN ACADÉMICA**: No simplificar conceptos complejos

### REGLAS ESPECÍFICAS:
1. **PRESERVAR ESTRUCTURA**: Mantener la organización y formato del contenido
2. **TERMINOLOGÍA CONSISTENTE**: Usar la misma traducción para términos técnicos repetidos
3. **CONTEXTO ACADÉMICO**: Considerar el nivel universitario en la elección de vocabulario
4. **FLUIDEZ NATURAL**: La traducción debe sonar natural en inglés americano

### VERIFICACIÓN OBLIGATORIA:
✅ ¿La traducción mantiene la precisión académica?
✅ ¿Usa terminología técnica apropiada en inglés?
✅ ¿Suena natural para hablantes de inglés americano?
✅ ¿Preserva el significado conceptual original?
✅ ¿Mantiene el nivel de formalidad universitaria?

Responde ÚNICAMENTE con el siguiente formato JSON:

{
  "ENUNCIADO (EN)": "Traducción precisa del enunciado al inglés americano",
  "CLAVE (EN)": "Traducción de la clave de respuesta (A, B, C, o D se mantienen igual)",
  "TEXTO A (EN)": "Traducción precisa de la opción A",
  "REALIMENTACIÓN A (EN)": "Traducción precisa de la realimentación A",
  "TEXTO B (EN)": "Traducción precisa de la opción B", 
  "REALIMENTACIÓN B (EN)": "Traducción precisa de la realimentación B",
  "TEXTO C (EN)": "Traducción precisa de la opción C",
  "REALIMENTACIÓN C (EN)": "Traducción precisa de la realimentación C",
  "TEXTO D (EN)": "Traducción precisa de la opción D",
  "REALIMENTACIÓN D (EN)": "Traducción precisa de la realimentación D"
}
`;

  try {
    let result = chatGPTClient.request(prompt_traduccion, "", null, null);
    console.log("Resultado traducción:", result);
    
    // Guardar los resultados en las columnas correspondientes
    sheet.saveResultToSheetRow(result, rowNumber, "ENUNCIADO (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "CLAVE (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "TEXTO A (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "REALIMENTACIÓN A (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "TEXTO B (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "REALIMENTACIÓN B (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "TEXTO C (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "REALIMENTACIÓN C (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "TEXTO D (EN)");
    sheet.saveResultToSheetRow(result, rowNumber, "REALIMENTACIÓN D (EN)");
    
  } catch (error) {
    console.error("Error traduciendo pregunta:", error);
    SpreadsheetApp.getUi().alert("Error al traducir pregunta: " + error.message);
  }
}

function aplicarEtiquetasNumericas(rowNumber, sheet) {
  try {
    // Obtener la clase de competencia
    const clase = sheet.getCellValue("Clase de Ítem", rowNumber) || "";
    const claseUpper = clase.toUpperCase().trim();
    const competencia = sheet.getCellValue("Competencia", rowNumber) || "";
    const competenciaUpper = competencia.toUpperCase().trim();
    console.log(`Aplicando etiquetas en fila ${rowNumber}: Clase="${claseUpper}", Competencia="${competenciaUpper}"`);
    // Mapeo de Competencias SABER
    const etiquetasBloom = {
      "INTERPRETATIVA": "1 Interpretativa",
      "ARGUMENTATIVA": "2 Argumentativa",
      "PROPOSITIVA": "3 Propositiva"
    };
    
    // Mapeo de Taxonomía de Bloom
    const etiquetasSaber = {
      "RECONOCIMIENTO": "1 Reconocimiento",
      "APLICACIÓN": "2 Aplicación",
      "ANÁLISIS": "3 Análisis",
      "EVALUACIÓN": "4 Evaluación",
      "SÍNTESIS": "5 Síntesis",
      "COMPRENSIÓN": "6 Comprensión"
    };
    
    // Aplicar etiquetas
    let etiquetaSaber = "";
    let etiquetaBloom = "";
    
    // Buscar en competencias SABER
    etiquetaSaber = etiquetasSaber[claseUpper];
    etiquetaBloom = etiquetasBloom[competenciaUpper];
    console.log(`Etiquetas encontradas: SABER="${etiquetaSaber}", Bloom="${etiquetaBloom}"`);
    // Si no se encontró en SABER pero sí en Bloom, aplicar solo Bloom
    // Si no se encontró en Bloom pero sí en SABER, aplicar solo SABER
    
    sheet.setCellValue("Etiqueta clase", rowNumber, etiquetaSaber);
    sheet.setCellValue("Etiqueta competencia", rowNumber, etiquetaBloom);
    
    
  } catch (error) {
    console.error(`Error aplicando etiquetas en fila ${rowNumber}:`, error);
    sheet.setCellValue("Etiqueta clase", rowNumber, "ERROR");
    sheet.setCellValue("Etiqueta competencia", rowNumber, "ERROR");
  }
}
