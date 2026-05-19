import { SPHttpClient } from '@microsoft/sp-http';
import { IPersonal } from '../models/IPersonal';

export class PersonalService {
    private _context: any;
    private _listName: string = "Personal EWS";

    constructor(context: any) { this._context = context; }

    public async getPersonal(): Promise<IPersonal[]> {
        try {
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items?$select=Id,Title,Rol,FotoPerfil,Email`;
            const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
                headers: { 'Accept': 'application/json;odata=nometadata', 'odata-version': '' }
            });

            if (!response.ok) return [];
            const data = await response.json();

            return (data.value || []).map((item: any) => ({
                Id: item.Id,
                NombreyApellido: item.Title,
                Rol: item.Rol,
                FotoPerfil: item.FotoPerfil ? item.FotoPerfil.Url : undefined,
                Email: item.Email
            }));
        } catch (error) {
            console.error("Error en getPersonal:", error);
            return [];
        }
    }

    /**
     * Obtiene los archivos de la biblioteca 'Fotos_Personal' para elegirlos en el formulario
     */
    public async getFotosDisponibles(): Promise<{ key: string, text: string, url: string }[]> {
        try {
            const serverRelativeUrl = `${this._context.pageContext.web.serverRelativeUrl}/Fotos_Personal`;
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files`;

            // Para listar archivos, odata=verbose suele ser más fiable y evita errores 406
            const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'odata-version': ''
                }
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error("Error al obtener archivos de la biblioteca:", errorText);
                return [];
            }

            const data = await response.json();
            // Con odata=verbose, los datos están en d.results
            const files = data.d && data.d.results ? data.d.results : [];

            return files.map((file: any) => ({
                key: `${window.location.origin}${file.ServerRelativeUrl}`,
                text: file.Name,
                url: `${window.location.origin}${file.ServerRelativeUrl}`
            }));
        } catch (error) {
            console.error("Error obteniendo fotos de la biblioteca:", error);
            return [];
        }
    }

    public async crearTrabajador(nuevo: { NombreyApellido: string, Rol: string, FotoPerfil?: string }): Promise<void> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;

        const body: any = {
            Title: nuevo.NombreyApellido,
            Rol: nuevo.Rol,
            FotoPerfil: nuevo.FotoPerfil ? {
                Description: nuevo.NombreyApellido,
                Url: nuevo.FotoPerfil
            } : null
        };

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=nometadata',
                'odata-version': '3.0'
            },
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            const err = await response.text();
            console.error("Detalle del error al crear ítem:", err);
            throw new Error("No se pudo crear el registro del personal.");
        }
    }

    public async getRolOptions(): Promise<string[]> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/fields?$filter=EntityPropertyName eq 'Rol'`;
        const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return [];
        const data = await response.json();
        return (data.value && data.value[0]) ? data.value[0].Choices : [];
    }

    public async actualizarTrabajador(id: number, datos: any): Promise<void> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
        const body = JSON.stringify({
            '__metadata': { 'type': `SP.Data.Personal_x0020_EWSListItem` },
            Title: datos.NombreyApellido,
            Rol: datos.Rol,
            FotoPerfil: datos.FotoPerfil ? {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Description': datos.NombreyApellido,
                'Url': datos.FotoPerfil
            } : null
        });

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            body: body,
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-type': 'application/json;odata=verbose',
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*',
                'odata-version': ''
            }
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Error detallado al actualizar:", errorText);

            // Si el error persiste por el nombre del tipo, probaremos una versión más simplificada
            throw new Error("No se pudo actualizar el registro del trabajador.");
        }
    }

    public async eliminarTrabajador(id: number): Promise<void> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;

        const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json',
                'X-HTTP-Method': 'DELETE',
                'IF-MATCH': '*'
            }
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("Error al eliminar de SharePoint:", errorText);
            throw new Error("No se pudo eliminar el registro.");
        }
    }
}