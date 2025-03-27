import { Component, OnInit } from '@angular/core';
import { ExcelProcessorService } from '../services/excel-processor.service';
import { TablaDetectada } from '../models/tabla-detectada.model';
import { NgClass, NgFor, NgIf } from '@angular/common';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';

@Component({
  selector: 'app-excel-to-xml',
  standalone: true,
  imports: [NgClass, NgIf, NgFor, ReactiveFormsModule, FormsModule],
  templateUrl: './excel-to-xml.component.html',
  styleUrls: ['./excel-to-xml.component.css'] // Cambiado de .scss a .css
})
export class ExcelToXmlComponent implements OnInit {
  xmlInput: string = '';
  xmlOutput: string = '';
  loading: boolean = false;
  processingExcel: boolean = false;
  error: string = '';
  tablasDetectadas: TablaDetectada[] = [];
  fileName: string = '';

  constructor(private excelProcessorService: ExcelProcessorService) { }

  ngOnInit(): void {
  }

  handleExcelUpload(event: any): void {
    const file = event.target.files[0];
    if (!file) return;
    
    this.fileName = file.name;
    this.processingExcel = true;
    this.error = '';
    
    const reader = new FileReader();
    
    reader.onload = (e: any) => {
      try {
        const data = new Uint8Array(e.target.result);
        
        // Convertir Excel a formato XML
        this.excelProcessorService.convertExcelToXml(data).then(xmlString => {
          this.xmlInput = xmlString;
          this.processingExcel = false;
        });
        
        // Opcionalmente, procesar automáticamente
        // this.limpiarXml(xmlString);
      } catch (err: any) {
        this.error = `Error al procesar el archivo Excel: ${err.message}`;
        console.error(err);
        this.processingExcel = false;
      }
    };
    
    reader.onerror = () => {
      this.error = "Error al leer el archivo";
      this.processingExcel = false;
    };
    
    reader.readAsArrayBuffer(file);
  }

  limpiarXml(xmlContent: string = this.xmlInput): void {
    this.loading = true;
    this.error = '';
    
    try {
      // Extraer las tablas del XML basadas en bordes
      const resultado = this.excelProcessorService.procesarExcelXmlConBordes(xmlContent);
      this.xmlOutput = resultado.xml;
      this.tablasDetectadas = resultado.tablas.map(t => ({
        id: t.id,
        tipo: t.tipo,
        encabezados: t.encabezados,
        filas: t.datos.length,
        columnas: t.encabezados.length,
        ubicacion: `Fila ${t.filaInicio+1}-${t.filaFin+1}, Columna ${t.columnaInicio}-${t.columnaFin}`
      } as TablaDetectada));
      
      this.loading = false;
    } catch (err: any) {
      this.error = `Error al procesar XML: ${err.message}`;
      console.error(err);
      this.loading = false;
    }
  }

  handleDownload(): void {
    if (!this.xmlOutput) return;
    
    const blob = new Blob([this.xmlOutput], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = this.fileName ? `${this.fileName.split('.')[0]}_clean.xml` : 'datos_limpios.xml';
    a.click();
    URL.revokeObjectURL(url);
  }
}