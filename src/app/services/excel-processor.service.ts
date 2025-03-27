import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import { TablaExcel } from '../models/tabla-excel.model';

@Injectable({
  providedIn: 'root',
})
export class ExcelProcessorService {
  constructor() {}

  async convertExcelToXml(data: Uint8Array): Promise<string> {
    const workbook = XLSX.read(data, { type: 'array' });
    return this.generateExcelXml(workbook);
  }

  private generateExcelXml(workbook: XLSX.WorkBook): string {
    // Obtener la primera hoja
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Obtener el rango de celdas
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

    // Crear el XML
    let xmlContent = '<?xml version="1.0"?>\n';
    xmlContent += '<?mso-application progid="Excel.Sheet"?>\n';
    xmlContent +=
      '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"\n';
    xmlContent += ' xmlns:o="urn:schemas-microsoft-com:office:office"\n';
    xmlContent += ' xmlns:x="urn:schemas-microsoft-com:office:excel"\n';
    xmlContent += ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"\n';
    xmlContent += ' xmlns:html="http://www.w3.org/TR/REC-html40">\n';

    // DocumentProperties
    xmlContent +=
      ' <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">\n';
    xmlContent += `  <Created>${new Date().toISOString()}</Created>\n`;
    xmlContent += `  <LastSaved>${new Date().toISOString()}</LastSaved>\n`;
    xmlContent += '  <Version>16.00</Version>\n';
    xmlContent += ' </DocumentProperties>\n';

    // Añadir algunas configuraciones estándar
    xmlContent +=
      ' <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">\n';
    xmlContent += '  <AllowPNG/>\n';
    xmlContent += ' </OfficeDocumentSettings>\n';

    // ExcelWorkbook
    xmlContent +=
      ' <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">\n';
    xmlContent += '  <WindowHeight>15720</WindowHeight>\n';
    xmlContent += '  <WindowWidth>29040</WindowWidth>\n';
    xmlContent += '  <WindowTopX>32767</WindowTopX>\n';
    xmlContent += '  <WindowTopY>32767</WindowTopY>\n';
    xmlContent += '  <ProtectStructure>False</ProtectStructure>\n';
    xmlContent += '  <ProtectWindows>False</ProtectWindows>\n';
    xmlContent += ' </ExcelWorkbook>\n';

    // Estilos
    xmlContent += ' <Styles>\n';
    xmlContent += '  <Style ss:ID="Default" ss:Name="Normal">\n';
    xmlContent += '   <Alignment ss:Vertical="Bottom"/>\n';
    xmlContent += '   <Borders/>\n';
    xmlContent +=
      '   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>\n';
    xmlContent += '   <Interior/>\n';
    xmlContent += '   <NumberFormat/>\n';
    xmlContent += '   <Protection/>\n';
    xmlContent += '  </Style>\n';
    xmlContent += '  <Style ss:ID="s17">\n';
    xmlContent += '   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/>\n';
    xmlContent += '   <Interior ss:Color="#D9D9D9" ss:Pattern="Solid"/>\n';
    xmlContent += '  </Style>\n';

    // Agregar estilo para celdas con bordes
    xmlContent += '  <Style ss:ID="s69">\n';
    xmlContent += '   <Borders>\n';
    xmlContent +=
      '    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>\n';
    xmlContent +=
      '    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>\n';
    xmlContent +=
      '    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>\n';
    xmlContent +=
      '    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>\n';
    xmlContent += '   </Borders>\n';
    xmlContent += '  </Style>\n';
    xmlContent += ' </Styles>\n';

    // Inicio de la hoja de trabajo
    xmlContent += ` <Worksheet ss:Name="${firstSheetName}">\n`;
    xmlContent += `  <Table ss:ExpandedColumnCount="${
      range.e.c + 1
    }" ss:ExpandedRowCount="${
      range.e.r + 1
    }" x:FullColumns="1" x:FullRows="1">\n`;

    // Columnas
    for (let C = 0; C <= range.e.c; ++C) {
      xmlContent += `   <Column ss:AutoFitWidth="0" ss:Width="80"/>\n`;
    }

    // Filas y celdas
    for (let R = 0; R <= range.e.r; ++R) {
      xmlContent += '   <Row>\n';

      for (let C = 0; C <= range.e.c; ++C) {
        const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
        const cell: XLSX.CellObject = worksheet[cell_address];

        if (cell) {
          // Identificar tipo de celda
          let cellType = 'String';
          let cellValue = '';

          switch (cell.t) {
            case 'n':
              cellType = 'Number';
              cellValue = String(cell.v);
              break;
            case 'b':
              cellType = 'Boolean';
              cellValue = cell.v ? '1' : '0';
              break;
            case 'd':
              cellType = 'DateTime';
              cellValue = cell.w || String(cell.v);
              break;
            default:
              cellType = 'String';
              cellValue = cell.v !== undefined ? String(cell.v) : '';

              // Escapar caracteres especiales XML
              cellValue = cellValue
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&apos;');
              break;
          }

          // Determinar estilo
          let styleID = '';

          // Primera fila con cabeceras
          if (R === 0) {
            styleID = ' ss:StyleID="s17"';
          }
          // Detectar si tiene bordes (simplificado - en realidad esto vendría de la hoja Excel)
          else if (cell.s && cell.s.border) {
            styleID = ' ss:StyleID="s69"';
          }

          xmlContent += `    <Cell${styleID}><Data ss:Type="${cellType}">${cellValue}</Data></Cell>\n`;
        } else {
          // Celda vacía
          xmlContent += '    <Cell/>\n';
        }
      }

      xmlContent += '   </Row>\n';
    }

    // Cierre de tabla y hoja
    xmlContent += '  </Table>\n';
    xmlContent +=
      '  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">\n';
    xmlContent += '   <PageSetup>\n';
    xmlContent += '    <Header x:Margin="0.3"/>\n';
    xmlContent += '    <Footer x:Margin="0.3"/>\n';
    xmlContent += '   </PageSetup>\n';
    xmlContent += '   <Selected/>\n';
    xmlContent += '   <ProtectObjects>False</ProtectObjects>\n';
    xmlContent += '   <ProtectScenarios>False</ProtectScenarios>\n';
    xmlContent += '  </WorksheetOptions>\n';
    xmlContent += ' </Worksheet>\n';
    xmlContent += '</Workbook>';

    return xmlContent;
  }

  procesarExcelXmlConBordes(xmlContent: string): {
    tablas: TablaExcel[];
    xml: string;
  } {
    // 1. Extraer información de estilos (bordes, cabeceras)
    const estilos = this.extraerInformacionEstilos(xmlContent);

    // 2. Extraer filas y celdas con información de estilos
    const filas = xmlContent.match(/<Row[\s\S]*?<\/Row>/g) || [];
    const gridDatos = filas.map((filaXml, idx) =>
      this.extraerCeldasConEstilos(filaXml, estilos)
    );

    // 3. Detectar tablas basadas en estructura y datos
    const tablas = this.detectarTablas(gridDatos);

    // 4. Generar XML
    return {
      tablas: tablas,
      xml: this.generarXmlDeTablas(tablas),
    };
  }

  private extraerInformacionEstilos(xmlContent: string): {
    [key: string]: any;
  } {
    const estilos: { [key: string]: any } = {};

    // Extraer definiciones de estilos
    const styleRegex = /<Style\s+ss:ID="([^"]+)"[^>]*>([\s\S]*?)<\/Style>/g;
    let matchStyle;

    while ((matchStyle = styleRegex.exec(xmlContent)) !== null) {
      const estiloId = matchStyle[1];
      const definicionEstilo = matchStyle[2];

      // Verificar si tiene bordes
      const tieneBordes =
        /<Borders>[\s\S]*?<\/Borders>/i.test(definicionEstilo) &&
        /<Border[\s\S]*?LineStyle[\s\S]*?>/i.test(definicionEstilo);

      // Verificar si es cabecera (negrita o fondo de color)
      const esCabecera =
        /<Font[^>]*\s+ss:Bold="1"/.test(definicionEstilo) ||
        /<Interior[^>]*\s+ss:Color=/.test(definicionEstilo);

      estilos[estiloId] = {
        id: estiloId,
        tieneBordes: tieneBordes,
        esCabecera: esCabecera,
        esNegrita: /<Font[^>]*\s+ss:Bold="1"/.test(definicionEstilo),
        tieneFondo: /<Interior[^>]*\s+ss:Color=/.test(definicionEstilo),
      };
    }

    return estilos;
  }

  private extraerCeldasConEstilos(
    filaXml: string,
    estilos: { [key: string]: any }
  ): any[] {
    const celdas: any[] = [];

    // Regex mejorado para capturar celdas con y sin datos
    const cellRegex =
      /<Cell(?:\s+ss:Index="(\d+)")?(?:\s+ss:StyleID="([^"]+)")?[^>]*>(?:<Data\s+ss:Type="[^"]*">([^<]*)<\/Data>)?<\/Cell>/g;
    let match;
    let indiceActual = 1; // Excel usa índices basados en 1

    while ((match = cellRegex.exec(filaXml)) !== null) {
      const indiceExplicito = match[1] ? parseInt(match[1]) : null;
      const estiloId = match[2] || null;
      const valor = match[3] || '';

      // Si hay un índice explícito, ajustamos la posición actual
      if (indiceExplicito) {
        indiceActual = indiceExplicito;
      }

      // Información de estilo
      const infoEstilo =
        estiloId && estilos[estiloId]
          ? estilos[estiloId]
          : {
              tieneBordes: false,
              esCabecera: false,
              esNegrita: false,
              tieneFondo: false,
            };

      // Guardamos el valor con su índice real y estilo
      celdas.push({
        indice: indiceActual,
        valor: valor,
        estiloId: estiloId,
        tieneBordes: infoEstilo.tieneBordes,
        esCabecera: infoEstilo.esCabecera,
        esNegrita: infoEstilo.esNegrita,
        tieneFondo: infoEstilo.tieneFondo,
      });

      indiceActual++;
    }

    return celdas;
  }

  private detectarTablasPorBordes(gridDatos: any[][]): TablaExcel[] {
    const tablas: TablaExcel[] = [];

    // Obtener todas las celdas con bordes
    const celdasConBordes: Array<{
      fila: number;
      columna: number;
      celda: any;
    }> = [];

    for (let i = 0; i < gridDatos.length; i++) {
      if (!gridDatos[i]) continue;

      for (const celda of gridDatos[i]) {
        // Verificar si la celda tiene bordes
        if (celda.tieneBordes || celda.estiloId === 's69') {
          celdasConBordes.push({
            fila: i,
            columna: celda.indice,
            celda: celda,
          });
        }
      }
    }

    // Si no hay celdas con bordes, no podemos detectar tablas por bordes
    if (celdasConBordes.length === 0) {
      console.warn('No se encontraron celdas con bordes en el documento.');
      return [];
    }

    // Agrupar celdas con bordes en regiones contiguas
    const celdasVisitadas: { [key: string]: boolean } = {};
    const regiones: Array<{
      filaInicio: number;
      filaFin: number;
      columnaInicio: number;
      columnaFin: number;
      celdas: Array<{ fila: number; columna: number; celda: any }>;
    }> = [];

    // Función para encontrar todas las celdas conectadas con bordes
    const encontrarRegionConBordes = (fila: number, columna: number) => {
      const clave = `${fila},${columna}`;
      if (celdasVisitadas[clave]) return null;

      celdasVisitadas[clave] = true;

      const celdaActual = celdasConBordes.find(
        (c) => c.fila === fila && c.columna === columna
      );
      if (!celdaActual) return null;

      // Iniciar una nueva región
      const region: {
        filaInicio: number;
        filaFin: number;
        columnaInicio: number;
        columnaFin: number;
        celdas: Array<{ fila: number; columna: number; celda: any }>;
      } = {
        filaInicio: fila,
        filaFin: fila,
        columnaInicio: columna,
        columnaFin: columna,
        celdas: [celdaActual],
      };

      // Cola para BFS (Breadth-First Search)
      const cola = [{ fila, columna }];

      while (cola.length > 0) {
        const { fila: f, columna: c } = cola.shift()!;

        // Comprobar celdas adyacentes (arriba, abajo, izquierda, derecha)
        const direcciones = [
          [-1, 0],
          [1, 0],
          [0, -1],
          [0, 1],
        ];

        for (const [df, dc] of direcciones) {
          const nuevaFila = f + df;
          const nuevaColumna = c + dc;
          const nuevaClave = `${nuevaFila},${nuevaColumna}`;

          // Verificar si esta nueva posición tiene una celda con borde y no ha sido visitada
          if (!celdasVisitadas[nuevaClave]) {
            const celdaAdyacente = celdasConBordes.find(
              (c) => c.fila === nuevaFila && c.columna === nuevaColumna
            );

            if (celdaAdyacente) {
              // Marcar como visitada
              celdasVisitadas[nuevaClave] = true;

              // Añadir a la región
              region.celdas.push(celdaAdyacente);

              // Actualizar límites de la región
              region.filaInicio = Math.min(region.filaInicio, nuevaFila);
              region.filaFin = Math.max(region.filaFin, nuevaFila);
              region.columnaInicio = Math.min(
                region.columnaInicio,
                nuevaColumna
              );
              region.columnaFin = Math.max(region.columnaFin, nuevaColumna);

              // Añadir a la cola para seguir explorando
              cola.push({ fila: nuevaFila, columna: nuevaColumna });
            }
          }
        }
      }

      // Solo considerar regiones de un tamaño mínimo (al menos 4 celdas)
      if (region.celdas.length >= 4) {
        return region;
      }

      return null;
    };

    // Buscar todas las regiones con bordes
    for (const { fila, columna } of celdasConBordes) {
      const region = encontrarRegionConBordes(fila, columna);
      if (region) {
        regiones.push(region);
      }
    }

    // Procesar cada región para extraer la tabla
    regiones.forEach((region, idx) => {
      // Verificar si esta región es parte de una región más grande ya detectada
      const esParte = regiones.some(
        (r, i) =>
          i !== idx &&
          region.filaInicio >= r.filaInicio &&
          region.filaFin <= r.filaFin &&
          region.columnaInicio >= r.columnaInicio &&
          region.columnaFin <= r.columnaFin
      );

      if (esParte) return; // Saltar si esta región es parte de otra más grande

      // Extraer encabezados (primera fila de la región)
      const filaEncabezado = region.filaInicio;
      const encabezados =
        gridDatos[filaEncabezado]
          ?.filter(
            (c) =>
              c.indice >= region.columnaInicio &&
              c.indice <= region.columnaFin &&
              c.valor !== ''
          )
          .sort((a, b) => a.indice - b.indice) || [];

      // Si no hay encabezados, no es una tabla válida
      if (encabezados.length === 0) return;

      // Extraer datos (filas restantes)
      const datos: string[][] = [];

      for (let i = filaEncabezado + 1; i <= region.filaFin; i++) {
        if (!gridDatos[i]) continue;

        const filaDatos: string[] = [];

        for (const encabezado of encabezados) {
          // Buscar la celda correspondiente a esta columna
          const celda = gridDatos[i].find(
            (c) => c.indice === encabezado.indice
          );
          filaDatos.push(celda ? celda.valor : '');
        }

        // Solo incluir filas que tengan al menos un valor no vacío
        if (filaDatos.some((v) => v !== '')) {
          datos.push(filaDatos);
        }
      }

      // Crear la tabla
      tablas.push({
        id: `tabla_${tablas.length + 1}`,
        tipo: tablas.length === 0 ? 'principal' : 'adicional',
        encabezados: encabezados.map((c) => c.valor),
        columnaInicio: region.columnaInicio,
        columnaFin: region.columnaFin,
        filaInicio: region.filaInicio,
        filaFin: region.filaFin,
        datos: datos,
      });
    });

    // Ordenar tablas por posición
    tablas.sort((a, b) => {
      if (a.filaInicio !== b.filaInicio) {
        return a.filaInicio - b.filaInicio;
      }
      return a.columnaInicio - b.columnaInicio;
    });

    // Renumerar tablas
    tablas.forEach((tabla, index) => {
      tabla.id = `tabla_${index + 1}`;
    });

    return tablas;
  }

  // Modificar el método principal para usar esta función o intentar ambos enfoques
  private detectarTablas(gridDatos: any[][]): TablaExcel[] {
    // Intentar primero detectar tablas basadas en bordes físicos
    const tablasPorBordes = this.detectarTablasPorBordes(gridDatos);

    // Si encontramos tablas con bordes, usarlas
    if (tablasPorBordes.length > 0) {
      return tablasPorBordes;
    }

    // Si no hay bordes en el documento, intentar detectar por agrupación de datos
    // Usaríamos aquí el otro algoritmo como respaldo
    console.warn(
      'No se detectaron tablas con bordes. Intentando detectar por agrupación de datos.'
    );

    // Este código continuaría con el algoritmo de detección por datos
    // que podemos implementar como respaldo

    // IMPORTANTE: Primero hay que verificar si es que el XML contiene información de bordes
    // El estilo s69 debería estar aplicado a celdas con bordes
    const hayEstilosDeBordes = gridDatos.some(
      (fila) =>
        fila &&
        fila.some((celda) => celda.tieneBordes || celda.estiloId === 's69')
    );

    if (!hayEstilosDeBordes) {
      console.warn(
        'El documento no parece contener información sobre bordes de celdas.'
      );
      // Aquí vendría el algoritmo alternativo
    }

    // Por ahora, si no encontramos tablas con bordes, devolvemos un array vacío
    return [];
  }
  
  private generarXmlDeTablas(tablas: TablaExcel[]): string {
    let xmlResult = '<?xml version="1.0" encoding="UTF-8"?>\n<datos>\n';

    tablas.forEach((tabla) => {
      // Usamos el ID tal como viene, que ahora será tabla_1, tabla_2, etc.
      xmlResult += `  <tabla id="${tabla.id}">\n`;

      // Metadatos de la tabla
      xmlResult += '    <metadatos>\n';
      xmlResult += `      <tipo>${tabla.tipo}</tipo>\n`;
      xmlResult += `      <fila_inicio>${tabla.filaInicio + 1}</fila_inicio>\n`; // +1 para mostrar índice humano
      xmlResult += `      <columna_inicio>${tabla.columnaInicio}</columna_inicio>\n`;
      xmlResult += '    </metadatos>\n';

      // Encabezados
      xmlResult += '    <encabezados>\n';
      tabla.encabezados.forEach((encabezado, i) => {
        xmlResult += `      <encabezado indice="${
          i + 1
        }">${encabezado}</encabezado>\n`;
      });
      xmlResult += '    </encabezados>\n';

      // Datos
      xmlResult += '    <filas>\n';
      tabla.datos.forEach((fila) => {
        xmlResult += '      <fila>\n';

        for (let i = 0; i < fila.length; i++) {
          if (fila[i] !== null && fila[i] !== '') {
            // Generar nombre seguro para la etiqueta XML basado en el encabezado
            const nombreColumna = tabla.encabezados[i]
              .replace(/\s+/g, '_')
              .replace(/[^a-zA-Z0-9_]/g, '')
              .toLowerCase();

            // Si el nombre generado está vacío o empieza con un número, añadir prefijo
            const nombreEtiqueta =
              nombreColumna === '' || /^\d/.test(nombreColumna)
                ? `columna_${i + 1}`
                : nombreColumna;

            // Escapar caracteres especiales en XML
            const valorEscapado = String(fila[i])
              .replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')
              .replace(/"/g, '&quot;')
              .replace(/'/g, '&apos;');

            xmlResult += `        <${nombreEtiqueta}>${valorEscapado}</${nombreEtiqueta}>\n`;
          }
        }

        xmlResult += '      </fila>\n';
      });
      xmlResult += '    </filas>\n';

      xmlResult += '  </tabla>\n';
    });

    xmlResult += '</datos>';

    return xmlResult;
  }
}
