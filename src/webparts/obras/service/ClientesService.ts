import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ICliente } from '../models/ICliente';

export class ClientesService {
    private _context: any;
    // Asegúrate de que este sea el nombre exacto de tu lista en SharePoint
    private _listName: string = "Clientes"; 

    constructor(context: any) { 
        this._context = context; 
    }

    /**
     * Obtiene todos los clientes de la lista
     */
    public async getClientes(): Promise<ICliente[]> {
        try {
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items?$select=Id,Title,CIF,Direccion,Email,Telefono`;
            const response: SPHttpClientResponse = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
                headers: { 
                    'Accept': 'application/json;odata=nometadata', 
                    'odata-version': '' 
                }
            });

            if (!response.ok) {
                console.error("Error al obtener clientes:", await response.text());
                return [];
            }

            const data = await response.json();

            // Mapeamos los datos de SharePoint a nuestra interfaz ICliente
            return (data.value || []).map((item: any) => ({
                Id: item.Id,
                Title: item.Title,
                CIF: item.CIF || "",
                Direccion: item.Direccion || "",
                Email: item.Email || "",
                Telefono: item.Telefono || ""
            }));
        } catch (error) {
            console.error("Excepción en getClientes:", error);
            return [];
        }
    }

    /**
     * Crea un nuevo cliente en la lista
     */
    public async crearCliente(nuevo: Partial<ICliente>): Promise<void> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;

        const body: any = {
            Title: nuevo.Title,
            CIF: nuevo.CIF,
            Direccion: nuevo.Direccion,
            Email: nuevo.Email,
            Telefono: nuevo.Telefono
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
            console.error("Detalle del error al crear cliente:", err);
            throw new Error("No se pudo crear el registro del cliente.");
        }
    }

    /**
     * Actualiza un cliente existente
     */
    public async actualizarCliente(id: number, datos: Partial<ICliente>): Promise<void> {
        const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
        
        // El tipo del metadata depende del nombre interno de tu lista. 
        // Si tu lista se llama "Clientes", suele ser "SP.Data.ClientesListItem".
        const body = JSON.stringify({
            '__metadata': { 'type': `SP.Data.${this._listName.replace(/ /g, '_x0020_')}ListItem` },
            Title: datos.Title,
            CIF: datos.CIF,
            Direccion: datos.Direccion,
            Email: datos.Email,
            Telefono: datos.Telefono
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
            console.error("Error detallado al actualizar cliente:", errorText);
            throw new Error("No se pudo actualizar el registro del cliente.");
        }
    }

    /**
     * Elimina un cliente por su ID
     */
    public async eliminarCliente(id: number): Promise<void> {
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
            console.error("Error al eliminar cliente de SharePoint:", errorText);
            throw new Error("No se pudo eliminar el registro.");
        }
    }
}