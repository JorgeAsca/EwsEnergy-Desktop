import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IObra } from "../models/IObra";
import { IObraCard } from "../models/IObraCard";
import { IFacepilePersona } from "@fluentui/react";

export class ProjectService {
  private _context: any;
  private _listName: string = "Proyectos y Obras";

  constructor(context: any) {
    this._context = context;
  }

  public async getObras(): Promise<IObra[]> {
    try {
      const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items?$select=Id,Title,Descripcion,DireccionObra,FechaInicio,FechaFinPrevista,EstadoObra,ProgresoReal,JornadasTotales,Cliente/Id,Cliente/Title&$expand=Cliente`;

      const response = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error("Error en la petición a SharePoint:", errorText);
        return [];
      }

      const data = await response.json();
      return data.value || [];
    } catch (error) {
      console.error("Error al obtener obras:", error);
      return [];
    }
  }

  public async crearObra(nuevaObra: any): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items`;
    const body = JSON.stringify({
      Title: nuevaObra.Nombre,
      ClienteId: nuevaObra.ClienteId,
      DireccionObra: nuevaObra.Direccion,
      FechaInicio: nuevaObra.FechaInicio,
      FechaFinPrevista: nuevaObra.FechaFin,
      JornadasTotales: nuevaObra.Jornadas,
      EstadoObra: "Fase Previa",
      ProgresoReal: 0,
    });

    await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: body,
      },
    );
  }

  public async updateObra(id: number, obraActualizada: any): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
    const body = JSON.stringify({
      Title: obraActualizada.Nombre,
      ClienteId: obraActualizada.ClienteId,
      DireccionObra: obraActualizada.Direccion,
      FechaInicio: obraActualizada.FechaInicio,
      FechaFinPrevista: obraActualizada.FechaFin,
      JornadasTotales: obraActualizada.Jornadas,
    });

    await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: body,
      },
    );
  }

  public async actualizarProgresoObra(
    id: number,
    nuevoProgreso: number,
  ): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
    const body = JSON.stringify({
      ProgresoReal: nuevoProgreso,
    });

    await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: body,
      },
    );
  }

  public async actualizarEstado(
    id: number,
    nuevoEstado: string,
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
        body: JSON.stringify({
          // Asegúrate de que 'Estado' sea el nombre interno de tu columna en SharePoint
          Estado: nuevoEstado,
        }),
      },
    );
  }

  public async getFotosPorObra(obraId: number): Promise<any[]> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Registro_Fotos_Diarias')/items?$filter=ObraId eq ${obraId}&$orderby=FechaRegistro desc`;

    try {
      const response = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) {
        // Esto nos imprimirá el error real de SharePoint en la consola
        const errorText = await response.text();
        console.error("Error detallado de SharePoint:", errorText);
        return [];
      }

      const data = await response.json();
      return data.value || [];
    } catch (e) {
      console.error("Error de red:", e);
      return [];
    }
  }

  public async getAsignacionesConPersonal(): Promise<any[]> {
    const siteUrl = this._context.pageContext.web.absoluteUrl;

    // 1. Obtenemos las asignaciones y expandimos el campo Personal (que debe ser un Lookup a la lista Personal_EWS)
    // Usamos $expand para traer los datos del operario en la misma consulta
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('Asignaciones_Obras')/items?$select=Id,ObraId,Personal/NombreyApellido,Personal/FotoPerfil&$expand=Personal`;

    try {
      const response = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error("Error al obtener asignaciones:", errorText);
        return [];
      }

      const data = await response.json();

      // Mapeamos los datos para que el componente Facepile los entienda fácilmente
      return (data.value || []).map((item: any) => ({
        Id: item.Id,
        ObraId: item.ObraId,
        Personal: {
          NombreyApellido: item.Personal
            ? item.Personal.NombreyApellido
            : "Sin nombre",
          FotoPerfil: item.Personal ? item.Personal.FotoPerfil : "",
        },
      }));
    } catch (error) {
      console.error("Error en getAsignacionesConPersonal:", error);
      return [];
    }
  }

  public async finalizarObra(id: number): Promise<void> {
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
        body: JSON.stringify({
          EstadoObra: "Finalizado",
        }),
      },
    );
  }

  public async cancelarObra(id: number): Promise<void> {
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
        body: JSON.stringify({
          EstadoObra: "Cancelado",
        }),
      },
    );
  }

  public async getObrasCompletas(
    asignaciones: any[],
    personal: any[],
  ): Promise<IObraCard[]> {
    const obras = await this.getObras();

    return obras.map((obra) => {
      // Filtrar operarios asignados a esta obra
      const asignados = asignaciones.filter(
        (a) => Number(a.ObraId) === Number(obra.Id),
      );
      const operariosProps: IFacepilePersona[] = asignados.map((asig) => {
        const p = personal.find(
          (pers) => Number(pers.Id) === Number(asig.PersonalId),
        );
        return { personaName: p ? p.NombreyApellido : "Desconocido" };
      });

      return {
        ...obra,
        clienteNombre: (obra as any).Cliente?.Title || "Sin Cliente",
        porcentajeReal: obra.ProgresoReal || 0,
        operarios: operariosProps,
        jornadasConsumidas: Math.round(
          ((obra.ProgresoReal || 0) / 100) * (obra.JornadasTotales || 30),
        ),
      } as IObraCard;
    });
  }
}
