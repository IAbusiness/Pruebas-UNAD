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

function reescribirPreguntasPorClase() {
  const sheet = new Formulation();
  const chatGPTClient = new ChatGPT();
  const selectedRows = sheet.getSelectedRows();
  
  if (selectedRows.length === 0) {
    SpreadsheetApp.getUi().alert("No hay filas válidas para procesar.");
    return;
  }
  
  // Crear columnas para la pregunta reescrita si no existen
  sheet.createColumn(["Pregunta Reescrita", "Opciones Reescritas", "Justificación Reescritura"]);
  
  for (const rowNumber of selectedRows) {
    console.log("Reescribiendo pregunta en fila:", rowNumber);
    reescribirPreguntaSegunClase(rowNumber, sheet, chatGPTClient);
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

function reescribirPreguntaSegunClase(rowNumber, sheet, chatGPTClient) {
  const data = {
    Temario: sheet.getSheet().getRange("B1").getValue(),
    Materia: sheet.getSheet().getRange("B2").getValue(),
    Clase: sheet.getCellValue("Clase de Ítem", rowNumber) || sheet.getCellValue("Clase", rowNumber),
    PreguntaOriginal: sheet.getCellValue("Pregunta", rowNumber) || sheet.getCellValue("Enunciado", rowNumber),
    OpcionA: sheet.getCellValue("Opción A", rowNumber) || sheet.getCellValue("A", rowNumber) || "",
    OpcionB: sheet.getCellValue("Opción B", rowNumber) || sheet.getCellValue("B", rowNumber) || "",
    OpcionC: sheet.getCellValue("Opción C", rowNumber) || sheet.getCellValue("C", rowNumber) || "",
    OpcionD: sheet.getCellValue("Opción D", rowNumber) || sheet.getCellValue("D", rowNumber) || "",
    ClaveRespuesta: sheet.getCellValue("Clave", rowNumber) || sheet.getCellValue("Clave de Respuesta", rowNumber),
    Competencia: sheet.getCellValue("Competencia", rowNumber) || ""
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

  const prompt_reescritura = `
Eres un experto en diseño de evaluaciones educativas. Tu tarea es REESCRIBIR una pregunta existente para que se ajuste perfectamente a las características de la clase de competencia especificada.

CONTEXTO EDUCATIVO:
- Materia: ${data.Materia}
- Temario: ${data.Temario}
- Competencia: ${data.Competencia}

CLASE DE COMPETENCIA OBJETIVO: ${claseSeleccionada}

CARACTERÍSTICAS DE ESTA CLASE:
- Descripción: ${caracteristicas.descripcion}
- Verbos típicos: ${caracteristicas.verbos.join(", ")}
- Enfoque: ${caracteristicas.enfoque}

PREGUNTA ORIGINAL:
${data.PreguntaOriginal}

OPCIONES ORIGINALES:
A) ${data.OpcionA}
B) ${data.OpcionB}  
C) ${data.OpcionC}
D) ${data.OpcionD}

RESPUESTA CORRECTA: ${data.ClaveRespuesta}

INSTRUCCIONES PARA LA REESCRITURA:

1. **MANTENER EL CONTENIDO TEMÁTICO**: La pregunta reescrita debe abordar el mismo tema/conocimiento que la original
2. **MANTENER LA RESPUESTA CORRECTA**: La opción que era correcta en la original debe seguir siendo correcta en la reescrita
3. **AJUSTAR AL TIPO DE COMPETENCIA**: Modifica la redacción y estructura para que requiera exactamente el tipo de proceso mental de la clase "${claseSeleccionada}"
4. **USAR VERBOS APROPIADOS**: Incorpora verbos típicos de esta competencia: ${caracteristicas.verbos.join(", ")}
5. **MANTENER RIGOR ACADÉMICO**: La pregunta debe ser clara, precisa y académicamente sólida
6. **OPCIONES BALANCEADAS**: Todas las opciones deben ser plausibles y estar al mismo nivel de dificultad

EJEMPLOS DE TRANSFORMACIÓN SEGÚN LA CLASE:
- Si es INTERPRETATIVA: Presenta información (texto, gráfico, caso) que debe ser interpretada
- Si es ARGUMENTATIVA: Requiere justificar, analizar o evaluar con fundamentos
- Si es PROPOSITIVA: Pide proponer soluciones o plantear hipótesis
- Si es RECONOCIMIENTO: Solicita recordar datos, definiciones o información factual
- Si es APLICACIÓN: Requiere usar conocimientos en situaciones específicas

Responde ÚNICAMENTE con el siguiente formato JSON:

{
  "Pregunta Reescrita": "Nueva versión de la pregunta adaptada a la clase ${claseSeleccionada}",
  "Texto A": "Opción A reescrita",
  "Texto B": "Opción B reescrita",
  "Texto C": "Opción C reescrita",
  "Texto D": "Opción D reescrita"
}
`;

  try {
    let result = chatGPTClient.request(prompt_reescritura, "", null, null);
    console.log("Resultado reescritura:", result);
    
    // Guardar los resultados en las columnas correspondientes
    sheet.saveResultToSheetRow(result, rowNumber, "Pregunta Reescrita");
    
    // Para las opciones, necesitamos procesarlas de manera especial
    const jsonString = result.copy
      .replace(/```json\n/, '')
      .replace(/```json/g, '')
      .replace(/```/g, '')
      .replace(/\n```/, '')
      .replace(/\n/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    
    const obj = JSON.parse(jsonString);
    
    // Formatear las opciones como texto
    const opcionesTexto = `A) ${obj["Opciones Reescritas"]["A"]}\nB) ${obj["Opciones Reescritas"]["B"]}\nC) ${obj["Opciones Reescritas"]["C"]}\nD) ${obj["Opciones Reescritas"]["D"]}`;
    sheet.setCellValue("Opciones Reescritas", rowNumber, opcionesTexto);
    
    sheet.saveResultToSheetRow(result, rowNumber, "Justificación Reescritura");
    
  } catch (error) {
    console.error("Error reescribiendo pregunta:", error);
    SpreadsheetApp.getUi().alert("Error al reescribir la pregunta: " + error.message);
  }
}
