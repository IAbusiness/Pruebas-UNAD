function reescribirPreguntasBasadoEnEvaluacion() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columnas para la pregunta mejorada si no existen
  sheet.createColumn([
    "Pregunta Mejorada", 
    "Texto A Mejorado", 
    "Texto B Mejorado", 
    "Texto C Mejorado", 
    "Texto D Mejorado",
    "Respuesta Correcta",
    "Justificación Mejora"
  ]);
  
  let preguntasProcesadas = 0;
  let preguntasOmitidas = 0;
  
  for (const rowNumber of selectedRows) {
    // Verificar si la pregunta necesita mejora (puntaje < 85)
    const puntajeTotal = sheet.getCellValue("Puntaje Total Evaluación", rowNumber) || 0;
    
    if (puntajeTotal < 85) {
      console.log(`Mejorando pregunta en fila ${rowNumber} (Puntaje: ${puntajeTotal}/100)`);
      mejorarPreguntaSegunEvaluacion(rowNumber, sheet, chatGPTClient);
      preguntasProcesadas++;
    } else {
      console.log(`Copiando pregunta original en fila ${rowNumber} (Puntaje: ${puntajeTotal}/100 - No requiere mejora)`);
      // Copiar contenido original a las columnas mejoradas
      copiarContenidoOriginal(rowNumber, sheet);
      preguntasOmitidas++;
    }
  }
  
  // Mostrar resumen del procesamiento
  const mensaje = `Procesamiento completado:\n- Preguntas mejoradas: ${preguntasProcesadas}\n- Preguntas copiadas (≥85 puntos): ${preguntasOmitidas}`;
  SpreadsheetApp.getUi().alert(mensaje);
  console.log(mensaje);
}

function copiarContenidoOriginal(rowNumber, sheet) {
  try {
    // Obtener contenido original
    const preguntaOriginal = sheet.getCellValue("Enunciado (con el contexto)", rowNumber) || "";
    const opcionA = sheet.getCellValue("Texto A (con el contexto)", rowNumber) || "";
    const opcionB = sheet.getCellValue("Texto B (con el contexto)", rowNumber) || "";
    const opcionC = sheet.getCellValue("Texto C (con el contexto)", rowNumber) || "";
    const opcionD = sheet.getCellValue("Texto D (con el contexto)", rowNumber) || "";
    const claveRespuesta = sheet.getCellValue("Clave", rowNumber) || sheet.getCellValue("Clave de Respuesta", rowNumber) || "";
    const puntajeTotal = sheet.getCellValue("Puntaje Total Evaluación", rowNumber) || 0;
    
    // Función para limpiar prefijos de las opciones
    function limpiarPrefijos(texto) {
      if (!texto) return "";
      
      // Eliminar prefijos como A), B), C), D) al inicio
      let textoLimpio = texto.replace(/^[A-D]\)\s*/i, "");
      
      // Eliminar prefijos como A., B., C., D. al inicio
      textoLimpio = textoLimpio.replace(/^[A-D]\.\s*/i, "");
      
      // Eliminar prefijos como A, B, C, D seguidos de espacio al inicio
      textoLimpio = textoLimpio.replace(/^[A-D]\s+/i, "");
      
      // Eliminar espacios extra al inicio y final
      return textoLimpio.trim();
    }
    
    // Limpiar prefijos de las opciones
    const opcionALimpia = limpiarPrefijos(opcionA);
    const opcionBLimpia = limpiarPrefijos(opcionB);
    const opcionCLimpia = limpiarPrefijos(opcionC);
    const opcionDLimpia = limpiarPrefijos(opcionD);
    
    // Copiar a las columnas mejoradas con contenido limpio
    sheet.setCellValue("Pregunta Mejorada", rowNumber, preguntaOriginal);
    sheet.setCellValue("Texto A Mejorado", rowNumber, opcionALimpia);
    sheet.setCellValue("Texto B Mejorado", rowNumber, opcionBLimpia);
    sheet.setCellValue("Texto C Mejorado", rowNumber, opcionCLimpia);
    sheet.setCellValue("Texto D Mejorado", rowNumber, opcionDLimpia);
    sheet.setCellValue("Respuesta Correcta", rowNumber, claveRespuesta);
    sheet.setCellValue("Justificación Mejora", rowNumber, `Contenido original conservado (prefijos eliminados) - Puntaje actual: ${puntajeTotal}/100 (≥85). La pregunta cumple con los estándares de calidad requeridos.`);
    
    console.log(`Contenido original copiado y limpio para fila ${rowNumber}`);
    console.log(`- Original A: "${opcionA}" → Limpio: "${opcionALimpia}"`);
    console.log(`- Original B: "${opcionB}" → Limpio: "${opcionBLimpia}"`);
    console.log(`- Original C: "${opcionC}" → Limpio: "${opcionCLimpia}"`);
    console.log(`- Original D: "${opcionD}" → Limpio: "${opcionDLimpia}"`);
    
  } catch (error) {
    console.error(`Error copiando contenido original en fila ${rowNumber}:`, error);
    sheet.setCellValue("Justificación Mejora", rowNumber, `Error al copiar contenido original: ${error.message}`);
  }
}

function mejorarPreguntaSegunEvaluacion(rowNumber, sheet, chatGPTClient) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    Clase: sheet.getCellValue("Clase de Ítem", rowNumber) || sheet.getCellValue("Clase", rowNumber),
    PreguntaOriginal: sheet.getCellValue("Enunciado (con el contexto)", rowNumber),
    OpcionA: sheet.getCellValue("Texto A (con el contexto)", rowNumber) || "",
    OpcionB: sheet.getCellValue("Texto B (con el contexto)", rowNumber) || "",
    OpcionC: sheet.getCellValue("Texto C (con el contexto)", rowNumber) || "",
    OpcionD: sheet.getCellValue("Texto D (con el contexto)", rowNumber) || "",
    ClaveRespuesta: sheet.getCellValue("Clave", rowNumber) || sheet.getCellValue("Clave de Respuesta", rowNumber),
    Competencia: sheet.getCellValue("Competencia", rowNumber) || "",
    
    // Datos de la evaluación previa
    EvaluacionTemario: sheet.getCellValue("Evaluación Temario", rowNumber) || 0,
    EvaluacionNivelUniversitario: sheet.getCellValue("Evaluación Nivel Universitario", rowNumber) || 0,
    EvaluacionConcordanciaClase: sheet.getCellValue("Evaluación Concordancia Clase", rowNumber) || 0,
    PuntajeTotal: sheet.getCellValue("Puntaje Total Evaluación", rowNumber) || 0,
    ObservacionesEvaluacion: sheet.getCellValue("Observaciones Evaluación", rowNumber) || ""
  };

  // Definir las características de cada clase
  const caracteristicasClases = {
    "INTERPRETATIVA": {
      descripcion: "Comprender o traducir el significado de un texto específico, de un problema, un esquema, una gráfica o un vídeo. Describir las variables involucradas en una situación problema y sus relaciones. Comprender posturas teóricas o planteamiento de Escuelas a partir de un texto o situación.",
      verbos: ["interpretar", "comprender", "explicar", "describir", "identificar", "reconocer", "traducir"],
      enfoque: "La pregunta debe requerir que el estudiante comprenda, interprete o traduzca información presentada en diferentes formas (texto, gráfico, esquema)."
    },
    "ARGUMENTATIVA": {
      descripcion: "Juzgar si un planteamiento es viable, dando razones. Analizar la ocurrencia de determinados hechos o situaciones con base en planteamientos teóricos. Seleccionar la mejor definición, argumentando la respuesta. Reconocer condiciones de necesidad o suficiencia.",
      verbos: ["justificar", "argumentar", "analizar", "evaluar", "juzgar", "sustentar", "demostrar"],
      enfoque: "La pregunta debe requerir que el estudiante analice, justifique o argumente con base en fundamentos teóricos o evidencia."
    },
    "PROPOSITIVA": {
      descripcion: "Resolver un problema o caso según condiciones determinadas por el contexto. Plantear hipótesis a partir de una situación o planteamiento.",
      verbos: ["proponer", "plantear", "resolver", "diseñar", "crear", "formular", "establecer"],
      enfoque: "La pregunta debe requerir que el estudiante proponga soluciones, plantee hipótesis o resuelva problemas de manera creativa."
    },
    "RECONOCIMIENTO": {
      descripcion: "Traer a la memoria datos, fechas, definiciones, enumeraciones, etc. Es el nivel más básico de proceso mental.",
      verbos: ["recordar", "identificar", "enumerar", "listar", "nombrar", "definir", "mencionar"],
      enfoque: "La pregunta debe requerir memoria directa de información factual, definiciones o datos específicos."
    },
    "COMPRENSIÓN": {
      descripcion: "Procesos de relación de información para interpretarla, desarrollando acciones como organización, discriminación, clasificación.",
      verbos: ["clasificar", "organizar", "discriminar", "relacionar", "comparar", "contrastar", "ejemplificar"],
      enfoque: "La pregunta debe requerir que el estudiante organice, clasifique o relacione información de manera estructurada."
    },
    "APLICACIÓN": {
      descripcion: "Seleccionar reglas, métodos, conceptos, principios, leyes o teorías para resolver una situación determinada.",
      verbos: ["aplicar", "utilizar", "emplear", "implementar", "ejecutar", "usar", "practicar"],
      enfoque: "La pregunta debe requerir que el estudiante aplique conocimientos teóricos a situaciones prácticas específicas."
    },
    "ANÁLISIS": {
      descripcion: "Identificar partes constitutivas de un todo, entender las relaciones entre ellas, comprender principios de organización y predecir causas y efectos.",
      verbos: ["analizar", "descomponer", "examinar", "distinguir", "diferenciar", "separar", "investigar"],
      enfoque: "La pregunta debe requerir que el estudiante descomponga información compleja en sus elementos constitutivos."
    },
    "EVALUACIÓN": {
      descripcion: "Formar juicios de valor frente a una situación, producto o procedimiento. Evidenciar capacidad predictiva y valorativa.",
      verbos: ["evaluar", "valorar", "criticar", "juzgar", "decidir", "seleccionar", "priorizar"],
      enfoque: "La pregunta debe requerir que el estudiante emita juicios fundamentados y tome decisiones basadas en criterios."
    },
    "SÍNTESIS": {
      descripcion: "Reunir elementos y formar un todo. Relacionada con la capacidad de producción intelectual y creación.",
      verbos: ["sintetizar", "integrar", "combinar", "componer", "crear", "construir", "elaborar"],
      enfoque: "La pregunta debe requerir que el estudiante combine elementos diversos para crear algo nuevo o una visión integral."
    }
  };

  const claseSeleccionada = data.Clase?.toUpperCase();
  const caracteristicas = caracteristicasClases[claseSeleccionada];

  if (!caracteristicas) {
    SpreadsheetApp.getUi().alert(`Clase "${data.Clase}" no reconocida. Clases válidas: ${Object.keys(caracteristicasClases).join(", ")}`);
    return;
  }

  const prompt_mejora = `
Eres un experto en diseño de evaluaciones educativas con amplia experiencia en mejora continua de preguntas de examen. Tu tarea es **MEJORAR*+ una pregunta existente basándote en una evaluación previa detallada que ya se realizó.

CONTEXTO EDUCATIVO:
- Materia: ${data.Materia}
- Temario: ${data.Temario}
- Competencia: ${data.Competencia}
- Clase de Competencia Objetivo: ${claseSeleccionada}

CARACTERÍSTICAS DE LA CLASE "${claseSeleccionada}":
- Descripción: ${caracteristicas.descripcion}
- Verbos típicos: ${caracteristicas.verbos.join(", ")}
- Enfoque: ${caracteristicas.enfoque}

PREGUNTA ACTUAL:
${data.PreguntaOriginal}

OPCIONES ACTUALES:
A) ${data.OpcionA}
B) ${data.OpcionB}
C) ${data.OpcionC}
D) ${data.OpcionD}

Respuesta Correcta: ${data.ClaveRespuesta}

## RESULTADOS DE LA EVALUACIÓN PREVIA:

### PUNTAJES OBTENIDOS:
- **Concordancia con Temario**: ${data.EvaluacionTemario}/100
- **Nivel Universitario**: ${data.EvaluacionNivelUniversitario}/100
- **Concordancia con Clase**: ${data.EvaluacionConcordanciaClase}/100
- **Puntaje Total**: ${data.PuntajeTotal}/100

### OBSERVACIONES DETALLADAS DE LA EVALUACIÓN:
${data.ObservacionesEvaluacion}

## INSTRUCCIONES PARA LA MEJORA:

### PRIORIDADES DE MEJORA (según puntajes más bajos):
1. **Si Temario < 80**: Asegurar que la pregunta esté específicamente alineada con "${data.Temario}" de la materia "${data.Materia}"
2. **Si Nivel Universitario < 80**: Elevar la complejidad, usar terminología técnica, requiere pensamiento crítico y análisis profundo
3. **Si Concordancia Clase < 80**: Ajustar para que requiera exactamente el proceso mental de "${claseSeleccionada}"

### ÁREAS ESPECÍFICAS A MEJORAR:
- **FORTALEZAS**: Mantener y potenciar los aspectos positivos identificados
- **DEBILIDADES**: Corregir específicamente los problemas señalados
- **RECOMENDACIONES**: Implementar las sugerencias concretas de la evaluación

### CRITERIOS DE MEJORA:
1. **Mantener coherencia temática**: La pregunta mejorada debe seguir abordando el mismo conocimiento específico
2. **Preservar la respuesta correcta**: El contenido que hace correcta la opción "${data.ClaveRespuesta}" debe mantenerse
3. **Implementar las observaciones**: Aplicar directamente las recomendaciones de la evaluación previa
4. **Elevar estándares**: Mejorar los aspectos con puntajes bajos según las observaciones específicas
5. **Optimizar redacción**: Usar verbos apropiados para la clase y terminología técnica universitaria

### REGLAS CRÍTICAS PARA LAS OPCIONES DE RESPUESTA:

⚠️ **OBLIGATORIO - EVITAR RESPUESTAS OBVIAS**:
- NUNCA crear opciones que sean claramente incorrectas o absurdas
- TODAS las opciones deben ser plausibles y creíbles para alguien que conoce parcialmente el tema
- Evitar respuestas que se puedan descartar fácilmente por sentido común
- Las opciones incorrectas deben representar errores conceptuales comunes o malentendidos típicos

⚠️ **OBLIGATORIO - FORMATO DE OPCIONES**:
- NO incluir prefijos como "A)", "B)", "C)", "D)" en el contenido de las opciones
- NO incluir prefijos como "A.", "B.", "C.", "D." en el contenido de las opciones
- Escribir ÚNICAMENTE el contenido de cada opción sin numeración o letras
- El sistema agregará automáticamente los identificadores

⚠️ **OBLIGATORIO - VALIDACIÓN DE RESPUESTA ÚNICA**:
- VERIFICAR que solo UNA opción sea completamente correcta
- ASEGURAR que las otras tres opciones sean definitivamente incorrectas
- REVISAR que no haya ambigüedad sobre cuál es la respuesta correcta
- Si hay matices, la opción correcta debe ser la MÁS completa y precisa

### ESTRATEGIAS ESPECÍFICAS SEGÚN LA CLASE:
- **INTERPRETATIVA**: Añadir contexto, gráficos, casos o información que requiera interpretación
- **ARGUMENTATIVA**: Incluir fundamentos teóricos que requieran análisis y justificación
- **PROPOSITIVA**: Presentar problemas o situaciones que requieran soluciones creativas
- **RECONOCIMIENTO**: Enfocar en datos, definiciones o información factual específica del temario
- **APLICACIÓN**: Crear escenarios prácticos donde se apliquen los conceptos teóricos

## VALIDACIÓN FINAL OBLIGATORIA:

Antes de generar tu respuesta, VERIFICA:
✅ ¿Todas las opciones son plausibles y creíbles?
✅ ¿Ninguna opción es obviamente incorrecta?
✅ ¿Solo UNA opción es completamente correcta?
✅ ¿Las opciones NO tienen prefijos A), B), C), D) o A., B., C., D.?
✅ ¿Las opciones incorrectas representan errores conceptuales realistas?

Responde ÚNICAMENTE con el siguiente formato JSON:

{
  "Pregunta Mejorada": "Versión mejorada de la pregunta que incorpora las observaciones de la evaluación y corrige las deficiencias identificadas",
  "Texto A Mejorado": "Opción A mejorada SIN prefijos como A) o A.",
  "Texto B Mejorado": "Opción B mejorada SIN prefijos como B) o B.", 
  "Texto C Mejorado": "Opción C mejorada SIN prefijos como C) o C.",
  "Texto D Mejorado": "Opción D mejorada SIN prefijos como D) o D.",
  "Respuesta Correcta": "${data.ClaveRespuesta}",
  "Justificación Mejora": "Explicación detallada de los cambios específicos realizados para atender las observaciones de la evaluación previa. Incluye cómo se abordaron las debilidades identificadas y se potenciaron las fortalezas. CONFIRMA que todas las opciones son plausibles y solo una es correcta."
}
`;

  try {
    let result = chatGPTClient.request(prompt_mejora, "", null, null);
    console.log("Resultado mejora basada en evaluación:", result);
    
    // Guardar los resultados en las columnas correspondientes
    sheet.saveResultToSheetRow(result, rowNumber, "Pregunta Mejorada");
    sheet.saveResultToSheetRow(result, rowNumber, "Texto A Mejorado");
    sheet.saveResultToSheetRow(result, rowNumber, "Texto B Mejorado");
    sheet.saveResultToSheetRow(result, rowNumber, "Texto C Mejorado");
    sheet.saveResultToSheetRow(result, rowNumber, "Texto D Mejorado");
    sheet.saveResultToSheetRow(result, rowNumber, "Respuesta Correcta");
    sheet.saveResultToSheetRow(result, rowNumber, "Justificación Mejora");
    
  } catch (error) {
    console.error("Error mejorando pregunta según evaluación:", error);
    SpreadsheetApp.getUi().alert("Error al mejorar la pregunta: " + error.message);
  }
}
