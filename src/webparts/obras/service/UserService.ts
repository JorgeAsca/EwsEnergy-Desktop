import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { RolUsuario } from '../models/IPersonal';

export class UserService {
    private _context: WebPartContext;
    private _baseUrl: string;

    constructor(context: WebPartContext) {
        this._context = context;
        this._baseUrl = context.pageContext.web.absoluteUrl;
    }

    /**
     * Determina el rol del usuario actual consultando sus grupos de SharePoint
     */
    public async getRolActual(): Promise<RolUsuario> {
        // Obtenemos los grupos del usuario actual
        const endpoint = `${this._baseUrl}/_api/web/currentuser/groups`;

        const response: SPHttpClientResponse = await this._context.spHttpClient.get(
            endpoint,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) {
            // Si hay error o no tiene grupos, por defecto es Operario (seguridad mínima)
            return 'Operario';
        }

        const data = await response.json();
        const grupos = data.value.map((g: any) => g.Title);

        if (grupos.indexOf('EWS_Admins') !== -1) return 'Administrador' as RolUsuario;
        if (grupos.indexOf('EWS_Managers') !== -1) return 'Manager' as RolUsuario;

        return 'Operario' as RolUsuario;
    }

    /**
     * Obtiene la información del perfil del usuario logueado
     */
    public getInfoUsuario() {
        return {
            nombre: this._context.pageContext.user.displayName,
            email: this._context.pageContext.user.email,
            id: this._context.pageContext.user.loginName
        };
    }
}