import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IObra } from "../models/IObra";
import { IObraCard } from "../models/IObraCard";
import { IFacepilePersona } from "@fluentui/react";

export class ProjectService {
  private _context: any;
  private _listName: string = "Proyectos y Obras";
  private _docLibraryName: string = "Documentos_Obras"; // <-- NOMBRE DE TU BIBLIOTECA

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

  // MODIFICADO: Ahora devuelve el ID (number) de la obra creada
  public async crearObra(nuevaObra: any): Promise<number> {
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

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json", // Sin nometadata para poder leer el ID devuelto
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "",
        },
        body: body,
      },
    );

    const data = await response.json();
    return data.Id; // Retorna el ID generado en SharePoint
  }

  // NUEVO: Convierte el File a buffer para subirlo
  private async _fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  // NUEVO: Crea la carpeta en la biblioteca si no existe
  public async asegurarCarpeta(nombreCarpeta: string): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/folders/add('${this._docLibraryName}/${nombreCarpeta}')`;
    try {
      await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json;odata=nometadata" }
      });
    } catch (e) {
      console.warn("La carpeta ya existe o hubo un aviso:", e);
    }
  }

  public async getDocumentosPorObra(idObra: number): Promise<any[]> {
    const nombreCarpeta = `Obra_${idObra}`;
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${this._docLibraryName}/${nombreCarpeta}')/Files`;
    
    try {
      const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
        headers: { Accept: "application/json" }
      });
      
      if (!response.ok) {
        return []; // Si da error (ej: la carpeta aún no existe), devolvemos array vacío
      }
      
      const data = await response.json();
      return data.value || [];
    } catch (e) {
      console.warn("No se encontraron documentos o la carpeta no existe:", e);
      return [];
    }
  }

  // NUEVO: Sube el archivo a la carpeta específica
  public async subirArchivoACarpeta(nombreCarpeta: string, file: File): Promise<void> {
    const buffer = await this._fileToArrayBuffer(file);
    // Limpiamos el nombre del archivo para que no de error en SharePoint por caracteres raros
    const cleanFileName = file.name.replace(/[^a-zA-Z0-9.\-_]/g, '_'); 
    
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${this._docLibraryName}/${nombreCarpeta}')/files/add(url='${cleanFileName}', overwrite=true)`;
    
    await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      headers: { Accept: "application/json;odata=nometadata" },
      body: buffer
    });
  }

  public async actualizarObra(id: number, obraActualizada: any): Promise<void> {
    return this.updateObra(id, obraActualizada);
  }

  public async eliminarObra(id: number): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
    await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json",
          "Content-type": "application/json",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
          "odata-version": "",
        },
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

  public async actualizarProgresoObra(id: number, nuevoProgreso: number): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._listName}')/items(${id})`;
    const body = JSON.stringify({ ProgresoReal: nuevoProgreso });

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

  public async actualizarEstado(id: number, nuevoEstado: string): Promise<void> {
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
          Estado: nuevoEstado,
        }),
      },
    );
  }

  public async getFotosPorObra(obraId: number): Promise<any[]> {
    // Añadido $expand=AttachmentFiles por si las fotos subidas son adjuntos
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Registro_Fotos_Diarias')/items?$filter=ObraId eq ${obraId}&$orderby=FechaRegistro desc&$expand=AttachmentFiles`;

    try {
      const response = await this._context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) {
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
      return (data.value || []).map((item: any) => ({
        Id: item.Id,
        ObraId: item.ObraId,
        Personal: {
          NombreyApellido: item.Personal ? item.Personal.NombreyApellido : "Sin nombre",
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
        body: JSON.stringify({ EstadoObra: "Finalizado" }),
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
        body: JSON.stringify({ EstadoObra: "Cancelado" }),
      },
    );
  }

  public async getObrasCompletas(asignaciones: any[], personal: any[]): Promise<IObraCard[]> {
    const obras = await this.getObras();

    return obras.map((obra) => {
      const asignados = asignaciones.filter((a) => Number(a.ObraId) === Number(obra.Id));
      const operariosProps: IFacepilePersona[] = asignados.map((asig) => {
        const p = personal.find((pers) => Number(pers.Id) === Number(asig.PersonalId));
        return { personaName: p ? p.NombreyApellido : "Desconocido" };
      });

      return {
        ...obra,
        clienteNombre: (obra as any).Cliente?.Title || "Sin Cliente",
        porcentajeReal: obra.ProgresoReal || 0,
        operarios: operariosProps,
        jornadasConsumidas: Math.round(((obra.ProgresoReal || 0) / 100) * (obra.JornadasTotales || 30)),
      } as IObraCard;
    });
  }

  /**
     * Obtiene los archivos guardados en la carpeta de la obra
     */
    public async getArchivosDeCarpeta(nombreCarpeta: string): Promise<any[]> {
        try {
            // CORREGIDO: Ahora busca en "Documentos_Obras" usando tu variable _docLibraryName
            const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${this._docLibraryName}/${nombreCarpeta}')/Files`;
            
            const response = await this._context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, {
                headers: { 'Accept': 'application/json;odata=nometadata' }
            });
            
            if (!response.ok) {
                console.warn("Carpeta vacía o no encontrada en SharePoint.");
                return [];
            }
            
            const data = await response.json();
            return data.value || [];
        } catch (e) {
            console.error("Error al obtener los archivos:", e);
            return [];
        }
    }

    /**
   * Elimina un archivo físico de la carpeta de la obra en SharePoint
   */
  public async eliminarArchivoDeCarpeta(nombreCarpeta: string, nombreArchivo: string): Promise<void> {
    // Construimos la URL relativa del archivo en SharePoint
    const serverRelativeUrl = `${this._context.pageContext.web.serverRelativeUrl}/${this._docLibraryName}/${nombreCarpeta}/${nombreArchivo}`;
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')`;

    const response = await this._context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json',
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error("No se pudo eliminar el archivo de SharePoint: " + errorText);
    }
  }
}