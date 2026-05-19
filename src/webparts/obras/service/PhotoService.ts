import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";

export class PhotoService {
  private _context: WebPartContext;
  private _libName: string = "Fotos_Diario";
  private _metadataListName: string = "Registro_Fotos_Diarias";

  constructor(context: WebPartContext) {
    this._context = context;
  }

  public async subirFotoProyecto(
    file: File,
    nombreProyecto: string,
    metadatos: {
      operario: string;
      operarioId: number;
      obraId: number;
      comentarios?: string;
    },
  ): Promise<void> {
    const archivoOptimizado = await this._comprimirImagen(file);
    const siteUrl = this._context.pageContext.web.absoluteUrl;
    const serverRelativeUrl = this._context.pageContext.web.serverRelativeUrl;

    // Limpiamos nombre de carpeta
    const nombreCarpeta = nombreProyecto.replace(/[/\\?%*:|"<>]/g, "-");
    const folderUrl = `${serverRelativeUrl}/${this._libName}/${nombreCarpeta}`;

    // 1. Asegurar carpeta
    await this._asegurarCarpeta(folderUrl);

    // 2. Subir archivo físico
    const fileName = `${Date.now()}_${metadatos.operarioId}_${encodeURIComponent(file.name)}`;
    const endpointFile = `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${folderUrl}')/files/add(url='${fileName}',overwrite=true)`;

    const uploadOptions: ISPHttpClientOptions = {
      body: file,
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-type": file.type,
        "odata-version": "3.0",
      },
    };

    const uploadResponse = await this._context.spHttpClient.post(
      endpointFile,
      SPHttpClient.configurations.v1,
      uploadOptions,
    );

    if (!uploadResponse.ok) throw new Error("Error al subir archivo.");

    const fileData = await uploadResponse.json();
    const fotoUrlAbsoluta = `${window.location.origin}${fileData.ServerRelativeUrl}`;

    // 3. Registrar metadatos vinculados al ObraId
    await this._registrarMetadatos(fotoUrlAbsoluta, nombreProyecto, metadatos);
  }

  private async _registrarMetadatos(
    url: string,
    proyecto: string,
    meta: any,
  ): Promise<void> {
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._metadataListName}')/items`;

    const body = {
      Title: proyecto,
      // Simplificamos el objeto URL para cumplir con la API de SharePoint
      UrlFoto: {
        Description: `Registro - ${proyecto}`,
        Url: url,
      },
      FechaRegistro: new Date().toISOString(),
      OperarioId: meta.operarioId,
      ObraId: meta.obraId,
      Comentarios: meta.comentarios || "",
    };

    const response = await this._context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify(body),
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-type": "application/json;odata=nometadata",
          "odata-version": "3.0",
        },
      },
    );

    if (!response.ok) {
      const err = await response.text();
      console.error("Error al registrar metadatos:", err);
      throw new Error("No se pudo crear el registro.");
    }
  }
  private async _asegurarCarpeta(folderUrl: string): Promise<void> {
    const siteUrl = this._context.pageContext.web.absoluteUrl;
    const checkEndpoint = `${siteUrl}/_api/web/getfolderbyserverrelativeurl('${folderUrl}')`;
    const checkResponse = await this._context.spHttpClient.get(
      checkEndpoint,
      SPHttpClient.configurations.v1,
    );

    if (checkResponse.status === 404) {
      const createEndpoint = `${siteUrl}/_api/web/folders`;
      await this._context.spHttpClient.post(
        createEndpoint,
        SPHttpClient.configurations.v1,
        {
          body: JSON.stringify({ ServerRelativeUrl: folderUrl }),
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=nometadata",
            "odata-version": "3.0",
          },
        },
      );
    }
  }
  public async getFotosHoyPorOperario(operarioId: number): Promise<any[]> {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this._metadataListName}')/items?$filter=OperarioId eq ${operarioId} and FechaRegistro ge '${hoy.toISOString()}'&$orderby=FechaRegistro desc`;
    const response = await this._context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
    );
    if (!response.ok) return [];
    const data = await response.json();
    return data.value || [];
  }
  public async uploadCompressedPhoto(
    file: File,
    nombreProyecto: string,
    metadatos: {
      operario: string;
      operarioId: number;
      obraId: number;
      comentarios?: string;
      latitud?: number;
      longitud?: number;
    },
  ): Promise<void> {
    // 1. Comprimimos
    const compressedFile = await this._comprimirImagen(file);

    // 2. Llamamos al método original con los 3 argumentos requeridos
    return this.subirFotoProyecto(compressedFile, nombreProyecto, metadatos);
  }
  private _comprimirImagen(file: File): Promise<File> {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (event: any) => {
        const img = new Image();
        img.src = event.target.result;
        img.onload = () => {
          const canvas = document.createElement("canvas");
          // Definimos un ancho máximo de 1200px para que se vea bien pero pese poco
          const MAX_WIDTH = 1200;
          const scaleSize = MAX_WIDTH / img.width;
          canvas.width = MAX_WIDTH;
          canvas.height = img.height * scaleSize;

          const ctx = canvas.getContext("2d");
          ctx?.drawImage(img, 0, 0, canvas.width, canvas.height);

          canvas.toBlob(
            (blob) => {
              // Devolvemos un nuevo archivo JPEG con calidad al 70%
              resolve(new File([blob!], file.name, { type: "image/jpeg" }));
            },
            "image/jpeg",
            0.7,
          );
        };
      };
    });
  }

  public async obtenerUbicacion(): Promise<{ lat: number; lng: number;} | null> {
    return new Promise((resolve) => {
      if (!navigator.geolocation) resolve(null);
      navigator.geolocation.getCurrentPosition(
        (pos) =>
          resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
        () => resolve(null),
        { enableHighAccuracy: true, timeout: 5000 },
      );
    });
  }
}
