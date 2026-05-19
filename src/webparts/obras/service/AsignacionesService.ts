import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IAsignacion } from "../models/IAsignacion";
import { ProjectService } from "./ProjectService";
import { PersonalService } from "./PersonalService";
import { IPersonal } from "../models/IPersonal";

export class AsignacionesService {
  private _context: WebPartContext;
  private _listName = "Asignaciones EWS";

  constructor(context: WebPartContext) {
    this._context = context;
  }

  //Metodo para cargar los datos de una soloza vez
  public async getDatosPanel() {
    const projectService = new ProjectService(this._context);
    const personalService = new PersonalService(this._context);

    const [obras, personal, asignaciones] = await Promise.all([
      projectService.getObras(),
      personalService.getPersonal(),
      this.getAsignaciones(),
    ]);

    return { obras, personal, asignaciones };
  }

  public async getAsignaciones(): Promise<IAsignacion[]> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;
    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
    );
    const data = await response.json();
    return data.value || [];
  }

  // Creación encapsulada en el servicio
  public async crearAsignacion(
    obraId: number,
    personalId: number,
    fechaFin: Date,
  ): Promise<void> {
    const body = {
      Title: `Asignación Obra ${obraId}`,
      ObraId: obraId,
      PersonalId: personalId,
      FechaInicio: new Date().toISOString(),
      FechaFinPrevista: fechaFin.toISOString(),
      EstadoProgreso: 0,
    };

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "content-type": "application/json;odata=nometadata",
        "odata-version": "3.0",
      },
      body: JSON.stringify(body),
    };

    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;
    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      options,
    );

    if (!response.ok) throw new Error("Error al guardar en SharePoint");
  }

  public async asignarPersonal(asignacion: IAsignacion): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;
    const body = {
      Title: `Asignación Obra ${asignacion.ObraId}`,
      ObraId: asignacion.ObraId,
      PersonalId: asignacion.PersonalId,
      FechaInicio: asignacion.FechaInicio.toISOString(),
      FechaFinPrevista: asignacion.FechaFinPrevista.toISOString(),
      EstadoProgreso: asignacion.EstadoProgreso,
    };

    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": "application/json;odata=nometadata",
        "odata-version": "3.0",
      },
      body: JSON.stringify(body),
    };

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      options,
    );

    if (!response.ok) {
      const error = await response.text();
      console.error("Detalle del error:", error);
      throw new Error("Error al guardar en SharePoint");
    }
  }

  public async eliminarAsignacion(id: number): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Asignaciones EWS')/items(${id})`;

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json",
          "Content-type": "application/json",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
          "odata-version": "3.0",
        },
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error("Error detallado de SharePoint:", errorText);
      throw new Error(`No se pudo eliminar: ${response.statusText}`);
    }
  }

  public async actualizarAsignacion(
    id: number,
    datos: Partial<IAsignacion>,
  ): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
    await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json",
          "Content-type": "application/json",
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*",
          "odata-version": "",
        },
        body: JSON.stringify(datos),
      },
    );
  }
  public getCuadrillaSugerida(obraId: number, operarioId: number, asignaciones: any[], personal: IPersonal[]): IPersonal[] {
    const idsEnObra = asignaciones
        .filter(a => Number(a.ObraId) === Number(obraId))
        .map(a => Number(a.PersonalId));
    
    return personal.filter(p => idsEnObra.indexOf(Number(p.Id)) !== -1 && p.Id !== operarioId);
}
}
