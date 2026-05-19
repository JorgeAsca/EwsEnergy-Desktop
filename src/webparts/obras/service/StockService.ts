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
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items`;
        const response: SPHttpClientResponse = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) throw new Error("Error al obtener el inventario");
        const data = await response.json();
        
        return data.value.map((item: any) => ({
            Id: item.Id,
            Title: item.Title,
            Categoria: item.Categor_x00ed_a || "General",
            StockActual: item.StockActual || 0,
            StockMinimo: item.StockM_x00ed_nimo || 0
        } as IMaterial));
    }

    public async crearMaterial(material: any): Promise<void> {
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
    }

    public async actualizarMaterial(id: number, material: any): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${id})`;
        const headers = { ...this._getHeaders(), 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };
        
        const body = JSON.stringify({
            Title: material.Title,
            Categor_x00ed_a: material.Categoria,
            StockActual: material.StockActual,
            StockM_x00ed_nimo: material.StockMinimo
        });

        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { headers, body });
    }

    public async actualizarStock(materialId: number, nuevaCantidad: number): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${materialId})`;
        const headers = { ...this._getHeaders(), 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };
        
        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { 
            headers, 
            body: JSON.stringify({ StockActual: nuevaCantidad }) 
        });
    }

    public async eliminarMaterial(id: number): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Inventario de Materiales')/items(${id})`;
        const headers = { ...this._getHeaders(), 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' };
        await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { headers });
    }
}