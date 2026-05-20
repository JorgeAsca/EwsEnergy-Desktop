import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IMaterial } from '../models/IMaterial';

export class StockService {
    private _context: WebPartContext;
    private _baseUrl: string;

    constructor(context: WebPartContext) {
        this._context = context;
        this._baseUrl = context.pageContext.web.absoluteUrl;
    }

    private _getHeaders() {
        return {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
        };
    }

    public async getInventario(): Promise<IMaterial[]> {
        // Añadimos $expand=AttachmentFiles para que traiga la foto
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items?$select=Id,Title,Categor_x00ed_a,StockActual,StockM_x00ed_nimo,AttachmentFiles&$expand=AttachmentFiles`;
        const response: SPHttpClientResponse = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        
        if (!response.ok) throw new Error("Error al obtener el inventario");
        const data = await response.json();
        
        return data.value.map((item: any) => ({
            Id: item.Id,
            Title: item.Title,
            Categoria: item.Categor_x00ed_a || "General",
            StockActual: item.StockActual || 0,
            StockMinimo: item.StockM_x00ed_nimo || 0,
            AttachmentFiles: item.AttachmentFiles // Guardamos la foto
        } as any));
    }

    // --- Helpers para subir fotos ---
    private fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result as ArrayBuffer);
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    private async subirAdjunto(itemId: number, archivo: File): Promise<void> {
        const fileData = await this.fileToArrayBuffer(archivo);
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${itemId})/AttachmentFiles/add(FileName='${archivo.name}')`;

        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/octet-stream',
                'odata-version': ''
            },
            body: fileData
        });
    }

    // --- CRUD con soporte para archivos ---
    public async crearMaterial(material: any, archivo: File | null): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items`;
        
        const body = JSON.stringify({
            Title: material.Title,
            Categor_x00ed_a: material.Categoria,
            StockActual: material.StockActual,
            StockM_x00ed_nimo: material.StockMinimo 
        });

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { 
            headers: this._getHeaders(),
            body: body 
        });

        if (!response.ok) {
            const errorRaw = await response.text();
            throw new Error("Error en SharePoint: " + errorRaw);
        }

        const nuevoItem = await response.json();

        // Si hay foto, la adjuntamos al item creado
        if (archivo) {
            await this.subirAdjunto(nuevoItem.Id, archivo);
        }
    }

    public async editarMaterial(id: number, material: any, archivoNuevo: File | null): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${id})`;
        const headers = { 
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'X-HTTP-Method': 'MERGE', 
            'IF-MATCH': '*' 
        };
        
        const body = JSON.stringify({
            '__metadata': { 'type': `SP.Data.Inventario_x0020_de_x0020_MaterialesListItem` },
            Title: material.Title,
            Categor_x00ed_a: material.Categoria,
            StockActual: material.StockActual,
            StockM_x00ed_nimo: material.StockMinimo
        });

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { headers, body });
        
        if (!response.ok) {
            const errorRaw = await response.text();
            throw new Error("Error en SharePoint al actualizar: " + errorRaw);
        }

        // Si han seleccionado una foto nueva, la subimos
        if (archivoNuevo) {
            await this.subirAdjunto(id, archivoNuevo);
        }
    }

    public async eliminarMaterial(id: number): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${id})`;
        const headers = { ...this._getHeaders(), 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' };
        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { headers });
    }

    public async actualizarStock(materialId: number, nuevaCantidad: number): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${materialId})`;
        const headers = { 
            ...this._getHeaders(), 
            'X-HTTP-Method': 'MERGE', 
            'IF-MATCH': '*' 
        };
        
        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { 
            headers, 
            body: JSON.stringify({ StockActual: nuevaCantidad }) 
        });
    }
}