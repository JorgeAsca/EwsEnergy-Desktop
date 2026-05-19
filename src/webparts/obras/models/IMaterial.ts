export interface IMaterial {
    Id: number;
    Title: string;
    Categoria: string;
    StockActual: number;
    StockMinimo: number;
    FotoMaterial?: {
        Url: string;
    };
}