export interface IReporteHistorial {
    Id: number;
    Title: string;
    Comentarios: string;
    FechaRegistro: string;
    OperarioId: number;
    ObraId: number;
    UrlFoto?:{
        Url: string;
        Description?: string;
    }

}