export interface IMaterial {
    id: number;
    titulo: string;
    stockActual: number;
    categoria: string;
    fotoUrl?: string; // Para la gestión de las fotos
}