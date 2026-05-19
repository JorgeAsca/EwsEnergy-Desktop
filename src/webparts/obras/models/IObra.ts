export type EstadoPresupuesto = 'PRESUPUESTO' | 'ACEPTADO OK' | 'STOCK ALMACEN';
export type EstadoObra = 'En Proceso' | 'Finalizado' | 'Cancelado';

export interface IObra {
    Id: number;
    Title: string; 
    Descripcion?: string;
    DireccionObra?: string;
    FechaInicio?: string;   
    FechaFinPrevista?: string; 
    EstadoObra: EstadoObra;
    JornadasTotales?: number;       
    Cliente?: {
        Title: string;
    };
    ProgresoReal?: number;

}