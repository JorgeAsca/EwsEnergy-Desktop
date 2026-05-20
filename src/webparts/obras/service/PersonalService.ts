import { SPHttpClient } from '@microsoft/sp-http';
import { IPersonal } from '../models/IPersonal';

export class PersonalService {
    private _context: any;
    private _listName: string = "Personal EWS";
    private _libraryName: string = "Fotos_Personal"; // Nombre de la biblioteca donde guardaremos las fotos

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
     * Convierte un string en Base64 a un ArrayBuffer, necesario para subir archivos a SharePoint
     */
    private base64ToArrayBuffer(base64: string): ArrayBuffer {
        const binary_string = window.atob(base64);
        const len = binary_string.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binary_string.charCodeAt(i);
        }
        return bytes.buffer;
    }

    /**
     * Sube una imagen en formato Base64 a la biblioteca de SharePoint y retorna la URL
     */
    private async subirImagen(base64Data: string, nombreEmpleado: string): Promise<string> {
        // Si no es un base64, asumimos que ya es una URL válida y la devolvemos tal cual
        if (!base64Data.startsWith('data:image')) {
            return base64Data;
        }

        try {
            // Limpiamos el nombre para usarlo en el archivo y generamos un nombre único
            const safeName = nombreEmpleado.replace(/[^a-zA-Z0-9]/g, "_").toLowerCase();
            const fileName = `foto_${safeName}_${new Date().getTime()}.jpg`;

            // Extraemos solo la parte del contenido base64 (quitamos el "data:image/jpeg;base64,")
            const base64Content = base64Data.split(',')[1];
            const buffer = this.base64ToArrayBuffer(base64Content);

            const serverRelativeUrl = `${this._context.pageContext.web.serverRelativeUrl}/${this._libraryName}`;
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files/add(url='${fileName}',overwrite=true)`;

            const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-type': 'application/octet-stream', // Importante para archivos
                    'odata-version': ''
                },
                body: buffer
            });

            if (!response.ok) {
                const err = await response.text();
                console.error("Error al subir imagen:", err);
                throw new Error("No se pudo subir la imagen a la biblioteca.");
            }

            const data = await response.json();
            // Retornamos la URL absoluta de la imagen recién subida
            return `${window.location.origin}${data.d.ServerRelativeUrl}`;
        } catch (error) {
            console.error("Error en subirImagen:", error);
            throw error;
        }
    }

    public async crearTrabajador(nuevo: { NombreyApellido: string, Rol: string, FotoPerfil?: string }): Promise<void> {
        let imageUrl = "";

        // Si hay una foto seleccionada, la subimos primero
        if (nuevo.FotoPerfil) {
            imageUrl = await this.subirImagen(nuevo.FotoPerfil, nuevo.NombreyApellido);
        }

        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;

        const body: any = {
            Title: nuevo.NombreyApellido,
            Rol: nuevo.Rol,
            FotoPerfil: imageUrl ? {
                Description: nuevo.NombreyApellido,
                Url: imageUrl
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

    public async actualizarTrabajador(id: number, datos: any): Promise<void> {
        let imageUrl = datos.FotoPerfil;

        // Si detectamos que la foto es un Base64 nuevo (el usuario eligió otra), la subimos
        if (imageUrl && imageUrl.startsWith('data:image')) {
            imageUrl = await this.subirImagen(imageUrl, datos.NombreyApellido);
        }

        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
        const body = JSON.stringify({
            '__metadata': { 'type': `SP.Data.Personal_x0020_EWSListItem` },
            Title: datos.NombreyApellido,
            Rol: datos.Rol,
            FotoPerfil: imageUrl ? {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Description': datos.NombreyApellido,
                'Url': imageUrl
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
            throw new Error("No se pudo actualizar el registro del trabajador.");
        }
    }

    public async getRolOptions(): Promise<string[]> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/fields?$filter=EntityPropertyName eq 'Rol'`;
        const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) return [];
        const data = await response.json();
        return (data.value && data.value[0]) ? data.value[0].Choices : [];
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