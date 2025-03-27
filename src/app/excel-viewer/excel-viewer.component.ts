import { Component, OnInit, ViewEncapsulation, signal, computed, inject, HostListener, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import { TablaExcel } from '../models/tabla-excel.model';
import { ToastrService } from 'ngx-toastr';

@Component({
  selector: 'app-excel-viewer',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './excel-viewer.component.html',
  styleUrls: ['./excel-viewer.component.scss'],
  encapsulation: ViewEncapsulation.None
})
export class ExcelViewerComponent implements OnInit {
  // ViewChild for file input
  @ViewChild('fileInput') fileInput!: ElementRef;
  
  // Services
  private toastr = inject(ToastrService);
  
  // Properties for file upload
  fileName = signal<string>('');
  loading = signal<boolean>(false);
  error = signal<string>('');
  isDragging = signal<boolean>(false);
  
  // Properties for Excel display
  workbook = signal<XLSX.WorkBook | null>(null);
  activeSheet = signal<string>('');
  gridData = signal<any[][]>([]);
  
  // Properties for selection
  selectionStart = signal<{ row: number, col: number } | null>(null);
  selectionEnd = signal<{ row: number, col: number } | null>(null);
  isSelecting = signal<boolean>(false);
  
  // Extracted tables array
  tablasExtraidas = signal<TablaExcel[]>([]);
  
  // Sheet dimensions
  maxRow = signal<number>(0);
  maxCol = signal<number>(0);
  columnHeaders = signal<string[]>([]);
  
  // Math for templates
  Math = Math;
  
  // Global document events
  @HostListener('document:mouseup')
  onDocumentMouseUp() {
    if (this.isSelecting()) {
      this.endSelection();
    }
  }
  
  // Drag and drop file handling
  @HostListener('document:dragover', ['$event'])
  onDragOver(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    if (this.isDropZoneVisible()) {
      this.isDragging.set(true);
    }
  }
  
  @HostListener('document:dragleave', ['$event'])
  onDragLeave(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDragging.set(false);
  }
  
  @HostListener('document:drop', ['$event'])
  onDrop(event: DragEvent) {
    event.preventDefault();
    event.stopPropagation();
    this.isDragging.set(false);
    
    if (this.isDropZoneVisible() && !this.loading()) {
      const files = event.dataTransfer?.files;
      if (files && files.length > 0) {
        const file = files[0];
        if (this.isExcelFile(file)) {
          this.processFile(file);
        } else {
          this.toastr.warning('El archivo debe ser de tipo Excel (.xlsx, .xls)', 'Formato inválido');
        }
      }
    }
  }
  
  // Computed values
  totalTablasExtraidas = computed(() => this.tablasExtraidas().length);
  
  constructor() { }

  ngOnInit(): void {
    // Optional: Initialize any required data or settings
  }

  // Helper to check if drop zone is visible
  isDropZoneVisible(): boolean {
    return !this.workbook() || this.workbook() === null;
  }
  
  // Helper to validate Excel files
  isExcelFile(file: File): boolean {
    return file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
  }
  
  // File handling
  handleFileUpload(event: any): void {
    const file = event.target.files[0];
    if (!file) return;
    
    this.processFile(file);
  }
  
  processFile(file: File): void {
    this.fileName.set(file.name);
    this.loading.set(true);
    this.error.set('');
    
    const reader = new FileReader();
    
    reader.onload = (e: any) => {
      try {
        const data = new Uint8Array(e.target.result);
        this.processExcel(data);
      } catch (err: any) {
        this.error.set(`Error al procesar el archivo Excel: ${err.message}`);
        console.error(err);
        this.loading.set(false);
        this.toastr.error('No se pudo procesar el archivo', 'Error');
      }
    };
    
    reader.onerror = () => {
      this.error.set("Error al leer el archivo");
      this.loading.set(false);
      this.toastr.error('No se pudo leer el archivo', 'Error');
    };
    
    reader.readAsArrayBuffer(file);
  }
  
  // Reset file input
  resetFile(): void {
    this.fileName.set('');
    if (this.fileInput) {
      this.fileInput.nativeElement.value = '';
    }
  }

  processExcel(data: Uint8Array): void {
    // Read workbook with all style and format options
    const wb = XLSX.read(data, { 
      type: 'array',
      cellStyles: true,
      cellDates: true,
      cellNF: true
    });
    
    this.workbook.set(wb);
    
    if (wb.SheetNames.length > 0) {
      this.activeSheet.set(wb.SheetNames[0]);
      this.renderSheet(wb.SheetNames[0]);
      this.toastr.success('Archivo Excel cargado correctamente', 'Éxito');
    }
    
    this.loading.set(false);
  }

  renderSheet(sheetName: string): void {
    const wb = this.workbook();
    if (!wb) return;
    
    const worksheet = wb.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    
    this.maxRow.set(range.e.r);
    this.maxCol.set(range.e.c);
    
    // Generate column headers (A, B, C, ...)
    const headers: string[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      headers.push(XLSX.utils.encode_col(c));
    }
    this.columnHeaders.set(headers);
    
    // Convert worksheet to array of rows and cells for rendering
    const grid: any[][] = [];
    
    for (let r = 0; r <= range.e.r; r++) {
      const row: any[] = [];
      
      for (let c = 0; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellRef];
        
        // Determine if the cell has borders
        const hasBorder = cell && cell.s && cell.s.border;
        
        row.push({
          ref: cellRef,
          value: cell ? XLSX.utils.format_cell(cell) : '',
          row: r,
          col: c,
          hasBorder: hasBorder
        });
      }
      
      grid.push(row);
    }
    
    this.gridData.set(grid);
  }

  changeSheet(sheetName: string): void {
    this.activeSheet.set(sheetName);
    this.renderSheet(sheetName);
    this.clearSelection();
    this.toastr.info(`Hoja "${sheetName}" seleccionada`, 'Cambio de hoja');
  }

  // Cell selection methods
  startSelection(row: number, col: number): void {
    this.selectionStart.set({ row, col });
    this.selectionEnd.set({ row, col });
    this.isSelecting.set(true);
  }

  updateSelection(row: number, col: number): void {
    if (this.isSelecting()) {
      this.selectionEnd.set({ row, col });
    }
  }

  endSelection(): void {
    this.isSelecting.set(false);
  }

  isCellSelected(row: number, col: number): boolean {
    const start = this.selectionStart();
    const end = this.selectionEnd();
    
    if (!start || !end) return false;
    
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    
    return row >= minRow && row <= maxRow && col >= minCol && col <= maxCol;
  }

  clearSelection(): void {
    this.selectionStart.set(null);
    this.selectionEnd.set(null);
    this.isSelecting.set(false);
  }

  // Extract selected table
  extraerTablaSeleccionada(): void {
    const start = this.selectionStart();
    const end = this.selectionEnd();
    const wb = this.workbook();
    
    if (!start || !end || !wb) {
      this.toastr.warning('Debe seleccionar un rango de celdas primero', 'Sin selección');
      return;
    }
    
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    
    // Get headers (first row of selection)
    const encabezados: string[] = [];
    const grid = this.gridData();
    
    for (let c = minCol; c <= maxCol; c++) {
      const cellValue = grid[minRow][c].value;
      encabezados.push(cellValue !== '' ? cellValue : `Col${c - minCol + 1}`);
    }
    
    // Get data (rest of rows)
    const datos: string[][] = [];
    for (let r = minRow + 1; r <= maxRow; r++) {
      const fila: string[] = [];
      for (let c = minCol; c <= maxCol; c++) {
        fila.push(grid[r][c].value);
      }
      
      // Only include rows with at least one non-empty value
      if (fila.some(v => v !== '')) {
        datos.push(fila);
      }
    }
    
    // Verify if there's enough data
    if (datos.length === 0) {
      this.toastr.warning('La selección no contiene datos. Seleccione un rango con al menos una fila de datos debajo de los encabezados.', 'Sin datos');
      return;
    }
    
    // Create table object
    const existingTables = this.tablasExtraidas();
    const newTable: TablaExcel = {
      id: `tabla_${existingTables.length + 1}`,
      tipo: existingTables.length === 0 ? 'principal' : 'adicional',
      encabezados: encabezados,
      columnaInicio: minCol,
      columnaFin: maxCol,
      filaInicio: minRow,
      filaFin: maxRow,
      datos: datos
    };
    
    // Add to extracted tables list
    this.tablasExtraidas.update(tables => [...tables, newTable]);
    
    // Clear selection
    this.clearSelection();
    
    this.toastr.success(`Tabla "${newTable.id}" extraída correctamente con ${datos.length} filas y ${encabezados.length} columnas`, 'Tabla extraída');
  }

  // Delete table from list
  eliminarTabla(index: number): void {
    const tablaEliminada = this.tablasExtraidas()[index];
    
    this.tablasExtraidas.update(tables => {
      const newTables = [...tables];
      newTables.splice(index, 1);
      
      // Renumber remaining tables
      newTables.forEach((tabla, i) => {
        tabla.id = `tabla_${i + 1}`;
        tabla.tipo = i === 0 ? 'principal' : 'adicional';
      });
      
      return newTables;
    });
    
    this.toastr.info(`Tabla "${tablaEliminada.id}" eliminada`, 'Tabla eliminada');
  }

  // Move table up in the list
  moverTablaArriba(index: number): void {
    if (index <= 0) return;
    
    this.tablasExtraidas.update(tables => {
      const newTables = [...tables];
      
      // Swap with previous table
      const temp = newTables[index];
      newTables[index] = newTables[index - 1];
      newTables[index - 1] = temp;
      
      // Renumber
      newTables.forEach((tabla, i) => {
        tabla.id = `tabla_${i + 1}`;
        tabla.tipo = i === 0 ? 'principal' : 'adicional';
      });
      
      return newTables;
    });
    
    this.toastr.info('Tabla movida hacia arriba', 'Orden actualizado');
  }

  // Move table down in the list
  moverTablaAbajo(index: number): void {
    const currentTables = this.tablasExtraidas();
    if (index >= currentTables.length - 1) return;
    
    this.tablasExtraidas.update(tables => {
      const newTables = [...tables];
      
      // Swap with next table
      const temp = newTables[index];
      newTables[index] = newTables[index + 1];
      newTables[index + 1] = temp;
      
      // Renumber
      newTables.forEach((tabla, i) => {
        tabla.id = `tabla_${i + 1}`;
        tabla.tipo = i === 0 ? 'principal' : 'adicional';
      });
      
      return newTables;
    });
    
    this.toastr.info('Tabla movida hacia abajo', 'Orden actualizado');
  }

  // Generate XML with extracted tables
  generarXML(): string {
    let xmlResult = '<?xml version="1.0" encoding="UTF-8"?>\n<datos>\n';
    
    this.tablasExtraidas().forEach(tabla => {
      xmlResult += `  <tabla id="${tabla.id}">\n`;
      
      // Table metadata
      xmlResult += '    <metadatos>\n';
      xmlResult += `      <tipo>${tabla.tipo}</tipo>\n`;
      xmlResult += `      <fila_inicio>${tabla.filaInicio + 1}</fila_inicio>\n`; // +1 for human index
      xmlResult += `      <columna_inicio>${tabla.columnaInicio}</columna_inicio>\n`;
      xmlResult += '    </metadatos>\n';
      
      // Headers
      xmlResult += '    <encabezados>\n';
      tabla.encabezados.forEach((encabezado, i) => {
        xmlResult += `      <encabezado indice="${i + 1}">${encabezado}</encabezado>\n`;
      });
      xmlResult += '    </encabezados>\n';
      
      // Data
      xmlResult += '    <filas>\n';
      tabla.datos.forEach(fila => {
        xmlResult += '      <fila>\n';
        
        for (let i = 0; i < fila.length; i++) {
          if (fila[i] !== null && fila[i] !== '') {
            // Generate safe XML tag name
            const nombreColumna = tabla.encabezados[i]
              .replace(/\s+/g, '_')
              .replace(/[^a-zA-Z0-9_]/g, '')
              .toLowerCase();
            
            // Add prefix if empty or starts with a number
            const nombreEtiqueta = nombreColumna === '' || /^\d/.test(nombreColumna) 
              ? `columna_${i + 1}` 
              : nombreColumna;
            
            // Escape special characters in XML
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

  // Download the generated XML
  descargarXML(): void {
    if (this.tablasExtraidas().length === 0) {
      this.toastr.warning('No hay tablas extraídas para generar XML', 'Sin datos');
      return;
    }
    
    const xml = this.generarXML();
    const blob = new Blob([xml], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = this.fileName() 
      ? `${this.fileName().split('.')[0]}_tablas.xml` 
      : 'tablas_extraidas.xml';
    a.click();
    
    URL.revokeObjectURL(url);
    
    this.toastr.success('XML generado y descargado correctamente', 'Descarga completa');
  }
}