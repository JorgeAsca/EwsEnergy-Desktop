import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { IReporteHistorial } from "../models/IReporteHistorial";
import { IDiarioEntrada } from "../models/IDiarioEntrada";


export class DailyReportService {
    private _context: WebPartContext;
    private _baseUrl: string;
    private _metadataListName: string = "Registro_Fotos_Diarias";

    constructor(context: WebPartContext) {
        this._context = context;
        this._baseUrl = context.pageContext.web.absoluteUrl;
    }

    /**
     * Guarda el reporte diario vinculando el texto con las URLs de las fotos
     */
    public async guardarReporteDiario(reporte: IDiarioEntrada): Promise<void> {
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('Diario de Trabajo')/items`;

        const body = JSON.stringify({
            Title: `Reporte - Obra ${reporte.ObraId} - ${reporte.Fecha}`,
            ObraId: reporte.ObraId,
            Comentarios: reporte.Comentarios,
            FotosRelacionadas: reporte.FotosUrls.join('; ')
        });

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            body: body,
            headers: {
                "Accept": "application/json",
                "Content-type": "application/json"
            }
        });

        if (!response.ok) throw new Error("No se pudo guardar el reporte diario.");
    }

    public async getHistorialGlobal(): Promise<IReporteHistorial[]> {
        // Optimizamos con $select para no traer basura de SharePoint
        const campos = "Id,Title,Comentarios,FechaRegistro,OperarioId,ObraId,UrlFoto";
        const endpoint = `${this._baseUrl}/_api/web/lists/getbytitle('${this._metadataListName}')/items?$select=${campos}&$orderby=FechaRegistro desc`;

        try {
            const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            
            if (!response.ok) {
                throw new Error(`Error al obtener historial: ${response.statusText}`);
            }

            const data = await response.json();
            return data.value || [];
        } catch (error) {
            console.error("Error en DailyReportService:", error);
            throw error;
        }
    }

    public async getFotosPorObra(obraId: number): Promise<any[]> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Registro_Fotos_Diarias')/items?$filter=ObraId eq ${obraId}&$orderby=FechaRegistro desc`;

        const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return [];
        const data = await response.json();
        return data.value || [];
    }


}