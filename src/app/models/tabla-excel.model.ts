export interface TablaExcel {
  id: string;
  tipo: string;
  encabezados: string[];
  columnaInicio: number;
  columnaFin: number;
  filaInicio: number;
  filaFin: number;
  datos: string[][];
}