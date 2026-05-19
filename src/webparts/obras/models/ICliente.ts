export interface ICliente {
    Id: number;
    Title: string; // Nombre del cliente
    CIF?: string;
    EmpresaRelacionadaId?: number; // ID del Lookup
}