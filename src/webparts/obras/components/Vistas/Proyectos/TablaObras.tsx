import * as React from "react";
import {
  Stack,
  Text,
  PrimaryButton,
  TextField,
  DatePicker,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Separator,
  Facepile,
  PersonaSize,
  Icon,
  Image,
  ImageFit,
  Modal,
  IconButton,
  DefaultButton,
  IDatePickerStrings,
  DayOfWeek,
  Dialog,
  DialogType,
  DialogFooter
} from "@fluentui/react";
import { SPHttpClient } from "@microsoft/sp-http";
import { ProjectService } from "../../../service/ProjectService";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { ClientesService } from "../../../service/ClientesService";
import { IObraCard } from "../../../models/IObraCard";
import styles from "./TablaObras.module.scss";

// ─── LEAFLET ────────────────────────────────────────────────────────────────
import L from "leaflet";
// ────────────────────────────────────────────────────────────────────────────

const stringsEspanol: IDatePickerStrings = {
  months: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
  shortMonths: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
  days: ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'],
  shortDays: ['D', 'L', 'M', 'X', 'J', 'V', 'S'],
  goToToday: 'Ir a hoy',
  prevMonthAriaLabel: 'Mes anterior',
  nextMonthAriaLabel: 'Mes siguiente',
  prevYearAriaLabel: 'Año anterior',
  nextYearAriaLabel: 'Año siguiente',
  closeButtonAriaLabel: 'Cerrar',
  monthPickerHeaderAriaLabel: '{0}, seleccione para cambiar el año',
  yearPickerHeaderAriaLabel: '{0}, seleccione para cambiar el mes',
  isRequiredErrorMessage: 'Este campo es obligatorio.',
  invalidInputErrorMessage: 'Formato de fecha no válido.',
};

export const TablaObras: React.FC<{ context: any }> = (props) => {
  const [obras, setObras] = React.useState<IObraCard[]>([]);
  const [clientes, setClientes] = React.useState<IDropdownOption[]>([]);
  const [obraSeleccionada, setObraSeleccionada] = React.useState<IObraCard | null>(null);
  const [fotosObra, setFotosObra] = React.useState<any[]>([]);

  // Estados locales para los archivos (Bandeja temporal de subida)
  const [archivosLocales, setArchivosLocales] = React.useState<{ nombre: string; icono: string; file: File }[]>([]);
  const [fotosPreviasLocales, setFotosPreviasLocales] = React.useState<{ url: string; file: File }[]>([]);

  // Archivos reales ya traídos desde la nube de SharePoint
  const [archivosNube, setArchivosNube] = React.useState<any[]>([]);

  // Para controlar qué foto se está viendo en grande (Lightbox)
  const [fotoAmpliada, setFotoAmpliada] = React.useState<string | null>(null);

  const [loading, setLoading] = React.useState(true);
  const [loadingFotos, setLoadingFotos] = React.useState(false);
  const [isOpen, setIsOpen] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  
  const [uploadingFiles, setUploadingFiles] = React.useState(false);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [obraEditandoId, setObraEditandoId] = React.useState<number | null>(null);

  const [isNuevoClienteOpen, setIsNuevoClienteOpen] = React.useState(false);
  const [nuevoClienteNombre, setNuevoClienteNombre] = React.useState("");
  const [creandoCliente, setCreandoCliente] = React.useState(false);

  const inputArchivoRef = React.useRef<HTMLInputElement>(null);
  const inputFotoRef = React.useRef<HTMLInputElement>(null);

  const [nuevaObra, setNuevaObra] = React.useState({
    Nombre: "", Descripcion: "", ClienteId: 0, Direccion: "",
    FechaInicio: new Date(), FechaFin: new Date(), JornadasTotales: 30,
  });

  const mapRef = React.useRef<HTMLDivElement>(null);
  const mapInstanceRef = React.useRef<L.Map | null>(null);
  const markerRef = React.useRef<L.Marker | null>(null);

  const services = React.useMemo(() => ({
    project: new ProjectService(props.context),
    personal: new PersonalService(props.context),
    asig: new AsignacionesService(props.context),
    clientes: new ClientesService(props.context),
  }), [props.context]);

  const cargarClientes = async (): Promise<IDropdownOption[]> => {
    try {
      const dataClientes = await services.clientes.getClientes();
      const opciones = dataClientes.map((c) => ({ key: c.Id!, text: c.Title }));
      setClientes(opciones);
      return opciones;
    } catch (error) {
      console.error("Error al refrescar listado de clientes:", error);
      return [];
    }
  };

  const cargarTodo = async () => {
    try {
      setLoading(true);
      const [listaObras, listaAsignaciones, listaPersonal, opcionesClientes] = await Promise.all([
        services.project.getObras(),
        services.asig.getAsignaciones(),
        services.personal.getPersonal(),
        cargarClientes()
      ]);

      const obrasProcesadas: IObraCard[] = listaObras.map((o) => {
        const porcentajeReal = (o.ProgresoReal || 0) / 100;
        const asigsObra = (listaAsignaciones as any[]).filter(a => Number(a.ObraId) === Number(o.Id));
        const operariosAsignados = Array.from(new Set(asigsObra.map(a => Number(a.PersonalId))))
          .map(pid => {
            const pers = (listaPersonal as any[]).find(p => Number(p.Id) === pid);
            return { personaName: pers?.NombreyApellido || "Operario", imageUrl: pers?.FotoPerfil || "" };
          });
        return {
          ...o,
          clienteNombre: opcionesClientes.find(c => Number(c.key) === (o as any).Cliente?.Id)?.text || "Cliente no definido",
          porcentajeReal: Math.min(Math.max(porcentajeReal, 0), 1),
          operarios: operariosAsignados,
          jornadasConsumidas: parseFloat((porcentajeReal * (o.JornadasTotales || 30)).toFixed(1)),
        };
      });

      setObras(obrasProcesadas);
      if (obraSeleccionada) {
        const actualizada = obrasProcesadas.find(o => o.Id === obraSeleccionada.Id);
        if (actualizada) setObraSeleccionada(actualizada);
      }
    } catch (e) {
      console.error("Error en Dashboard:", e);
    } finally { setLoading(false); }
  };

  React.useEffect(() => { cargarTodo(); }, []);

  const handleCrearClienteRapido = async () => {
    if (!nuevoClienteNombre.trim()) return;
    try {
      setCreandoCliente(true);
      await services.clientes.crearCliente({ Title: nuevoClienteNombre });
      const nuevasOpciones = await cargarClientes();
      const clienteCreado = nuevasOpciones.find(c => c.text.toLowerCase() === nuevoClienteNombre.trim().toLowerCase());
      if (clienteCreado) {
        setNuevaObra(prev => ({ ...prev, ClienteId: clienteCreado.key as number }));
      }
      setIsNuevoClienteOpen(false);
      setNuevoClienteNombre("");
    } catch (e) {
      alert("Error al dar de alta al cliente exprés.");
    } finally {
      setCreandoCliente(false);
    }
  };

  React.useEffect(() => {
    if (!isOpen) return;
    const timer = setTimeout(() => {
      if (!mapRef.current || mapInstanceRef.current) return;
      if (!document.getElementById("leaflet-css")) {
        const link = document.createElement("link");
        link.id = "leaflet-css";
        link.rel = "stylesheet";
        link.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
        document.head.appendChild(link);
      }

      delete (L.Icon.Default.prototype as any)._getIconUrl;
      L.Icon.Default.mergeOptions({
        iconRetinaUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
        iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
        shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
      });

      const map = L.map(mapRef.current, { tap: false, dragging: true } as L.MapOptions).setView([40.416775, -3.70379], 6);

      if (mapRef.current) {
        L.DomEvent.disableClickPropagation(mapRef.current);
        L.DomEvent.disableScrollPropagation(mapRef.current);
      }

      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        attribution: '© OpenStreetMap', maxZoom: 19,
      }).addTo(map);

      setTimeout(() => map.invalidateSize(), 100);

      map.on("click", (e: L.LeafletMouseEvent) => {
        const { lat, lng } = e.latlng;
        if (markerRef.current) markerRef.current.setLatLng([lat, lng]);
        else markerRef.current = L.marker([lat, lng], { draggable: true }).addTo(map);
      });
      mapInstanceRef.current = map;
    }, 400); 
    
    return () => {
      clearTimeout(timer);
      if (mapInstanceRef.current) {
        mapInstanceRef.current.remove();
        mapInstanceRef.current = null;
        markerRef.current = null;
      }
    };
  }, [isOpen]);

  React.useEffect(() => {
    if (!nuevaObra.Direccion || nuevaObra.Direccion.length < 5) return;
    const debounce = setTimeout(async () => {
      try {
        const resp = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(nuevaObra.Direccion)}&limit=1`, { headers: { "Accept-Language": "es" } });
        const data = await resp.json();
        if (data.length > 0 && mapInstanceRef.current) {
          const coords: L.LatLngExpression = [parseFloat(data[0].lat), parseFloat(data[0].lon)];
          mapInstanceRef.current.flyTo(coords, 16, { animate: true, duration: 1.2 });
          if (markerRef.current) markerRef.current.setLatLng(coords);
          else markerRef.current = L.marker(coords, { draggable: true }).addTo(mapInstanceRef.current);
          markerRef.current.bindPopup(`<b>📍 ${nuevaObra.Direccion}</b>`).openPopup();
        }
      } catch (e) { console.warn("Error geocodificando:", e); }
    }, 800);
    return () => clearTimeout(debounce);
  }, [nuevaObra.Direccion]);

  const refrescarArchivosNube = async (idObra: number) => {
    if (services.project.getArchivosDeCarpeta) {
      const nombreCarpeta = `Obra_${idObra}`;
      const archivos = await services.project.getArchivosDeCarpeta(nombreCarpeta);
      setArchivosNube(archivos || []);
    }
  };

  const verDetallesObra = async (obra: IObraCard) => {
    setObraSeleccionada(obra);
    setLoadingFotos(true);
    setArchivosLocales([]); 
    setFotosPreviasLocales([]);
    setArchivosNube([]); 

    try {
      const fotos = await services.project.getFotosPorObra(obra.Id as number);
      setFotosObra(fotos || []);
      await refrescarArchivosNube(obra.Id as number);
    } catch (e) { console.error(e); } 
    finally { setLoadingFotos(false); }
  };

  const handleSubirPlano = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      let icono = "Document";
      if (file.name.endsWith(".pdf")) icono = "PDF";
      if (file.name.endsWith(".dwg")) icono = "VisioDocument";
      if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) icono = "ExcelDocument";
      setArchivosLocales(prev => [...prev, { nombre: file.name, icono, file }]);
      if (inputArchivoRef.current) inputArchivoRef.current.value = ""; 
    }
  };

  const handleSubirFoto = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        setFotosPreviasLocales(prev => [...prev, { url: reader.result as string, file: file }]);
        if (inputFotoRef.current) inputFotoRef.current.value = ""; 
      };
      reader.readAsDataURL(file);
    }
  };

  const eliminarArchivoLocal = (index: number) => {
    setArchivosLocales(prev => prev.filter((_, i) => i !== index));
  };

  const eliminarFotoLocal = (index: number) => {
    setFotosPreviasLocales(prev => prev.filter((_, i) => i !== index));
  };

  const handleConfirmarSubidaArchivos = async () => {
    if (!obraSeleccionada) return;
    try {
      setUploadingFiles(true);
      const idObra = obraSeleccionada.Id as number;
      const nombreCarpeta = `Obra_${idObra}`;
      
      const archivosParaSubir = [
        ...archivosLocales.map(a => a.file),
        ...fotosPreviasLocales.map(f => f.file)
      ];

      if (archivosParaSubir.length > 0) {
        await services.project.asegurarCarpeta(nombreCarpeta);
        for (const file of archivosParaSubir) {
          await services.project.subirArchivoACarpeta(nombreCarpeta, file);
        }
        alert("✅ ¡Todos los archivos e imágenes se han subido a SharePoint!");
        setArchivosLocales([]);
        setFotosPreviasLocales([]);
        await refrescarArchivosNube(idObra);
      }
    } catch (error) {
      alert("❌ Ocurrió un error al subir los archivos. Por favor, inténtalo de nuevo.");
      console.error("Error subiendo archivos:", error);
    } finally {
      setUploadingFiles(false);
    }
  };

  const handleGuardar = async () => {
    if (!nuevaObra.Nombre || !nuevaObra.ClienteId) return;
    try {
      setSaving(true);
      if (obraEditandoId) {
        await services.project.updateObra(obraEditandoId, nuevaObra);
      } else {
        await services.project.crearObra(nuevaObra);
      }
      setIsOpen(false);
      resetForm();
      await cargarTodo();
    } catch (e) { 
      alert("Error al guardar los datos de la obra.");
      console.error(e);
    } 
    finally { setSaving(false); }
  };

  const resetForm = () => {
    setObraEditandoId(null);
    setNuevaObra({ Nombre: "", Descripcion: "", ClienteId: 0, Direccion: "", FechaInicio: new Date(), FechaFin: new Date(), JornadasTotales: 30 });
  };

  const handleAccionObra = async (id: number, accion: "finalizar" | "cancelar" | "eliminar") => {
    const confirmacion = { finalizar: "¿Estás seguro de finalizar esta obra?", cancelar: "¿Deseas cancelar esta obra?", eliminar: "⚠️ ¿ESTÁS SEGURO? Se borrarán todos los registros permanentemente." };
    if (!window.confirm(confirmacion[accion])) return;
    try {
      setIsProcessing(true);
      if (accion === "finalizar") await services.project.finalizarObra(id);
      if (accion === "cancelar") await services.project.cancelarObra(id);
      if (accion === "eliminar") {
        await props.context.spHttpClient.post(`${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Proyectos y Obras')/items(${id})`, SPHttpClient.configurations.v1, {
          headers: { Accept: "application/json", "IF-MATCH": "*", "X-HTTP-Method": "DELETE" },
        });
      }
      setObraSeleccionada(null);
      await cargarTodo();
    } catch (e) { console.error(e); } 
    finally { setIsProcessing(false); }
  };

  const renderProgressTracker = (pReal: number) => {
    const totalBoxes = 10;
    const filledBoxes = Math.round(pReal * totalBoxes);
    return (
      <div className={styles.progressTrackerBox} title={`Avance: ${(pReal * 100).toFixed(0)}%`}>
        {Array.from({ length: totalBoxes }).map((_, idx) => (
          <div key={idx} className={`${styles.trackerDot} ${idx < filledBoxes ? styles.filledOnTrack : ""}`} />
        ))}
      </div>
    );
  };

  const handleEliminarArchivoNube = async (nombreArchivo: string) => {
    if (!obraSeleccionada) return;
    if (!window.confirm(`¿Estás seguro de que deseas eliminar permanentemente el archivo "${nombreArchivo}" de SharePoint?`)) return;

    try {
      setLoadingFotos(true); 
      const idObra = obraSeleccionada.Id as number;
      const nombreCarpeta = `Obra_${idObra}`;
      await services.project.eliminarArchivoDeCarpeta(nombreCarpeta, nombreArchivo);
      alert("✅ Archivo eliminado correctamente de SharePoint.");
      await refrescarArchivosNube(idObra);
    } catch (error) {
      alert("❌ No se pudo eliminar el archivo. Inténtalo de nuevo.");
      console.error(error);
    } finally {
      setLoadingFotos(false);
    }
  };

  const obrasAgrupadas = obras.reduce((acc, obra) => {
    const estado = obra.EstadoObra || "Sin Asignar";
    if (!acc[estado]) acc[estado] = [];
    acc[estado].push(obra);
    return acc;
  }, {} as Record<string, IObraCard[]>);

  const planosNube = archivosNube.filter(a => !a.Name.match(/\.(jpg|jpeg|png|gif|webp)$/i));
  const fotosNube = archivosNube.filter(a => a.Name.match(/\.(jpg|jpeg|png|gif|webp)$/i));

  if (loading && obras.length === 0) return <Spinner size={SpinnerSize.large} label="Sincronizando Dashboard EWS..." />;

  return (
    <div className={styles.container}>
      <div className={styles.headerSection}>
        <Stack>
          <Text variant="xxLarge" className={styles.tituloPrincipal}>Panel de Control de Obras</Text>
          <Text variant="small" className={styles.subtituloHeader}>Gestión y seguimiento EWS Energy</Text>
        </Stack>
        <PrimaryButton iconProps={{ iconName: "Add" }} text="Nueva Obra" onClick={() => { resetForm(); setIsOpen(true); }} className={styles.btnNuevaObra} />
      </div>

      <div className={styles.splitLayout}>
        <div className={styles.listColumn}>
          <div className={styles.listContainer}>
            {Object.keys(obrasAgrupadas).length === 0 && <MessageBar>No hay proyectos registrados.</MessageBar>}
            {Object.keys(obrasAgrupadas).map((estado) => (
              <div key={estado}>
                <Text className={styles.listGroupHeader}>{estado}</Text>
                {obrasAgrupadas[estado].map((o) => (
                  <div key={o.Id} className={`${styles.listItem} ${obraSeleccionada?.Id === o.Id ? styles.selected : ""}`} onClick={() => verDetallesObra(o)}>
                    <Text className={styles.obraTitle}>{o.Title}</Text>
                    {renderProgressTracker(o.porcentajeReal)}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </div>

        <div className={styles.detailColumn}>
          {obraSeleccionada ? (
            <div className={styles.detailContent}>
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack>
                  <Text variant="xLarge" className={styles.detailTitle}>{obraSeleccionada.Title}</Text>
                  <Text variant="small" style={{ color: "#666" }}>{obraSeleccionada.clienteNombre}</Text>
                </Stack>
                <div className={`${styles.badgeEstado} ${obraSeleccionada.EstadoObra === "Finalizado" ? styles.finalizado : obraSeleccionada.EstadoObra === "Cancelado" ? styles.cancelado : styles.activo}`}>
                  {obraSeleccionada.EstadoObra || "Fase Previa"}
                </div>
                {/* CORREGIDO AQUÍ: Al pinchar en editar, cargamos la variable correcta "JornadasTotales" de la obra seleccionada */}
                <DefaultButton iconProps={{ iconName: "Edit" }} text="Editar" onClick={() => {
                  setObraEditandoId(obraSeleccionada.Id as number);
                  setNuevaObra({ 
                    Nombre: obraSeleccionada.Title, 
                    Descripcion: obraSeleccionada.Descripcion || "", 
                    ClienteId: (clientes.find(c => c.text === obraSeleccionada.clienteNombre)?.key as number) || 0, 
                    Direccion: obraSeleccionada.DireccionObra || "", 
                    FechaInicio: new Date(obraSeleccionada.FechaInicio || Date.now()), 
                    FechaFin: new Date(obraSeleccionada.FechaFinPrevista || Date.now()), 
                    JornadasTotales: obraSeleccionada.JornadasTotales || 30 
                  });
                  setIsOpen(true);
                }} />
              </Stack>
              <Separator />
              <Stack horizontal tokens={{ childrenGap: 40 }} className={styles.infoSection}>
                <Stack><Text className={styles.labelSeccion}>Dirección</Text><Text><Icon iconName="MapPin" className={styles.iconVerde} /> {obraSeleccionada.DireccionObra || "Sin dirección"}</Text></Stack>
                <Stack><Text className={styles.labelSeccion}>Jornadas Consumidas</Text><Text><Icon iconName="Calendar" className={styles.iconVerde} /> {obraSeleccionada.jornadasConsumidas} / {obraSeleccionada.JornadasTotales || 30}</Text></Stack>
                <Stack><Text className={styles.labelSeccion}>Avance Físico</Text><Text><Icon iconName="CompletedSolid" className={styles.iconVerde} /> {(obraSeleccionada.porcentajeReal * 100).toFixed(0)}% Ejecutado</Text></Stack>
                <Stack><Text className={styles.labelSeccion}>Equipo en Campo</Text>
                  {obraSeleccionada.operarios?.length > 0 ? <Facepile personas={obraSeleccionada.operarios} personaSize={PersonaSize.size32} /> : <Text variant="small" style={{ fontStyle: "italic", color: "#888" }}>Sin personal</Text>}
                </Stack>
              </Stack>

              <div className={styles.planosSection}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 15 } }}>
                  <Text variant="large" className={styles.sectionTitle}>Planos y Documentación</Text>
                  <DefaultButton iconProps={{ iconName: "Upload" }} className={styles.btnUpload} onClick={() => inputArchivoRef.current?.click()}>Añadir Archivo</DefaultButton>
                  <input type="file" ref={inputArchivoRef} style={{ display: "none" }} onChange={handleSubirPlano} />
                </Stack>
                
                <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
                  {planosNube.map((archivo, idx) => {
                    let icono = "Document";
                    if (archivo.Name.endsWith(".pdf")) icono = "PDF";
                    if (archivo.Name.endsWith(".dwg")) icono = "VisioDocument";
                    if (archivo.Name.endsWith(".xlsx") || archivo.Name.endsWith(".xls")) icono = "ExcelDocument";
                    
                    const tenantUrl = props.context.pageContext.web.absoluteUrl.replace(props.context.pageContext.web.serverRelativeUrl, "");
                    const link = (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1") ? tenantUrl + archivo.ServerRelativeUrl : archivo.ServerRelativeUrl;

                    return (
                      <div key={`nube-plano-${idx}`} className={styles.planoCard}>
                        <Icon iconName={icono} className={icono === 'PDF' ? styles.pdfIcon : styles.dwgIcon} />
                        <Text variant="smallPlus">
                          <a href={encodeURI(link)} target="_blank" rel="noreferrer" style={{ textDecoration: 'none', color: 'inherit', fontWeight: 600 }}>
                            {archivo.Name}
                          </a>
                        </Text>
                        <IconButton
                          iconProps={{ iconName: "Cancel" }}
                          title="Eliminar archivo permanentemente de SharePoint"
                          onClick={() => handleEliminarArchivoNube(archivo.Name)}
                          styles={{ root: { position: 'absolute', right: 2, top: '50%', transform: 'translateY(-50%)', width: 24, height: 24 } }}
                        />
                      </div>
                    )
                  })}

                  {archivosLocales.map((archivo, idx) => (
                    <div key={`local-plano-${idx}`} className={styles.planoCard} style={{ position: 'relative', paddingRight: '28px', border: '1px dashed #0078d4' }}>
                      <Icon iconName={archivo.icono} className={archivo.icono === 'PDF' ? styles.pdfIcon : styles.dwgIcon} />
                      <Text variant="smallPlus">{archivo.nombre} <i style={{fontSize: 10, color: '#0078d4'}}>(Cola)</i></Text>
                      <IconButton
                        iconProps={{ iconName: "Cancel" }}
                        title="Quitar de la cola"
                        onClick={() => eliminarArchivoLocal(idx)}
                        styles={{ root: { position: 'absolute', right: 2, top: '50%', transform: 'translateY(-50%)', width: 24, height: 24 } }}
                      />
                    </div>
                  ))}

                  {planosNube.length === 0 && archivosLocales.length === 0 && (
                    <Text variant="small" style={{ fontStyle: "italic", color: "#888" }}>No hay documentos adjuntos en esta obra.</Text>
                  )}
                </Stack>
              </div>

              <div className={styles.planosSection} style={{ marginTop: 20 }}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 15 } }}>
                  <Text variant="large" className={styles.sectionTitle}>Fotografías Previas / Finales</Text>
                  <DefaultButton iconProps={{ iconName: "Camera" }} className={styles.btnUpload} onClick={() => inputFotoRef.current?.click()}>Añadir Foto</DefaultButton>
                  <input type="file" accept="image/*" ref={inputFotoRef} style={{ display: "none" }} onChange={handleSubirFoto} />
                </Stack>

                <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                  {fotosNube.map((foto, idx) => {
                    const tenantUrl = props.context.pageContext.web.absoluteUrl.replace(props.context.pageContext.web.serverRelativeUrl, "");
                    const link = (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1") ? tenantUrl + foto.ServerRelativeUrl : foto.ServerRelativeUrl;
                    const cleanLink = encodeURI(link);
                    
                    return (
                      <div key={`nube-foto-${idx}`} style={{ position: "relative", display: "inline-block" }}>
                        <Image 
                          src={cleanLink} 
                          width={150} 
                          height={100} 
                          imageFit={ImageFit.cover} 
                          style={{ borderRadius: '6px', border: '1px solid #ccc', cursor: 'pointer' }} 
                          onClick={() => setFotoAmpliada(cleanLink)}
                        />
                        <IconButton
                          iconProps={{ iconName: "Cancel" }}
                          title="Eliminar foto permanentemente de SharePoint"
                          onClick={() => handleEliminarArchivoNube(foto.Name)}
                          styles={{
                            root: { position: 'absolute', top: 4, right: 4, backgroundColor: 'rgba(255,255,255,0.9)', borderRadius: '50%', width: 22, height: 22 },
                            icon: { fontSize: 10, color: '#d13438', fontWeight: 'bold' }
                          }}
                        />
                      </div>
                    );
                  })}

                  {fotosPreviasLocales.map((fotoObj, idx) => (
                    <div key={`local-foto-${idx}`} style={{ position: "relative", display: "inline-block" }}>
                      <Image 
                        src={fotoObj.url} 
                        width={150} 
                        height={100} 
                        imageFit={ImageFit.cover} 
                        style={{ borderRadius: '6px', border: '2px dashed #0078d4', opacity: 0.7, cursor: 'pointer' }} 
                        onClick={() => setFotoAmpliada(fotoObj.url)}
                      />
                      <IconButton
                        iconProps={{ iconName: "Cancel" }}
                        title="Quitar de la cola"
                        onClick={() => eliminarFotoLocal(idx)}
                        styles={{
                          root: { position: 'absolute', top: 4, right: 4, backgroundColor: 'rgba(255,255,255,0.9)', borderRadius: '50%', width: 22, height: 22 },
                          icon: { fontSize: 10, color: '#d13438', fontWeight: 'bold' }
                        }}
                      />
                    </div>
                  ))}

                  {fotosNube.length === 0 && fotosPreviasLocales.length === 0 && (
                    <Text variant="small" style={{ fontStyle: "italic", color: "#888" }}>No hay fotos previas cargadas.</Text>
                  )}
                </Stack>
              </div>

              {(archivosLocales.length > 0 || fotosPreviasLocales.length > 0) && (
                <div style={{ marginTop: 25, padding: '15px 20px', backgroundColor: '#f3f2f1', borderRadius: '8px', border: '1px solid #0078d4' }}>
                  <Stack horizontal verticalAlign="center" horizontalAlign="space-between">
                    <Stack>
                      <Text variant="mediumPlus" style={{ fontWeight: 600, color: '#0078d4' }}>
                        🚀 Tienes {archivosLocales.length + fotosPreviasLocales.length} archivo(s) pendientes de subir
                      </Text>
                      <Text variant="small" style={{ color: '#605e5c' }}>
                        Haz clic en el botón para guardarlos definitivamente en la nube del proyecto.
                      </Text>
                    </Stack>
                    <PrimaryButton 
                      text={uploadingFiles ? "Guardando en nube..." : "Guardar en Obra"} 
                      iconProps={{ iconName: "CloudUpload" }} 
                      onClick={handleConfirmarSubidaArchivos} 
                      disabled={uploadingFiles} 
                    />
                  </Stack>
                </div>
              )}

              <div className={styles.historialSection}>
                <Text variant="large" className={styles.sectionTitle}>Reportes de Jornada</Text>
                {loadingFotos ? <Spinner size={SpinnerSize.large} label="Cargando reportes..." /> : fotosObra.length > 0 ? (
                  <Stack tokens={{ childrenGap: 15 }} styles={{ root: { marginTop: 15 } }}>
                    {fotosObra.map((f, i) => {
                      let fotoSrc = "";
                      
                      if (f.AttachmentFiles && f.AttachmentFiles.length > 0) {
                        fotoSrc = f.AttachmentFiles[0].ServerRelativeUrl;
                      } else if (typeof f.UrlFoto === 'string' && f.UrlFoto.startsWith('{')) {
                        try { const imgData = JSON.parse(f.UrlFoto); fotoSrc = imgData.serverRelativeUrl; } catch(e){}
                      } else if (f.UrlFoto && f.UrlFoto.Url) {
                        fotoSrc = f.UrlFoto.Url;
                      }

                      if (fotoSrc && (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1")) {
                        const tenantUrl = props.context.pageContext.web.absoluteUrl.replace(props.context.pageContext.web.serverRelativeUrl, "");
                        fotoSrc = tenantUrl + fotoSrc;
                      }

                      const cleanFotoSrc = fotoSrc ? encodeURI(fotoSrc) : "";

                      return (
                        <div key={i} className={styles.fotoCard}>
                          <Stack horizontal tokens={{ childrenGap: 15 }}>
                            {cleanFotoSrc ? (
                              <Image 
                                src={cleanFotoSrc} 
                                width={120} 
                                height={90} 
                                imageFit={ImageFit.cover} 
                                className={styles.fotoThumb} 
                                style={{ borderRadius: '4px', border: '1px solid #edebe9', cursor: 'pointer' }}
                                onClick={() => setFotoAmpliada(cleanFotoSrc)} 
                              />
                            ) : (
                              <div style={{ width: 120, height: 90, backgroundColor: '#f3f2f1', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid #edebe9' }}>
                                <Icon iconName="Photo2" style={{ fontSize: 24, color: '#c8c6c4' }} />
                              </div>
                            )}
                            <Stack>
                              <Text className={styles.fotoFecha}>📅 {new Date(f.FechaRegistro).toLocaleDateString()} - Worker {f.Operario}</Text>
                              <div className={styles.fotoComentarioBox}><Text className={styles.fotoComentarioText}>"{f.Comentarios || "Sin observaciones técnicas"}"</Text></div>
                            </Stack>
                          </Stack>
                        </div>
                      );
                    })}
                  </Stack>
                ) : <MessageBar messageBarType={MessageBarType.info}>No hay reportes diarios para esta obra.</MessageBar>}
              </div>

              <div className={styles.planosSection}>
                <Separator />
                <Stack tokens={{ childrenGap: 15 }} style={{ marginTop: "20px" }}>
                  <Text variant="large" className={styles.sectionTitle}>Gestión de Obra</Text>
                  <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                    {isProcessing ? <Spinner label="Procesando..." /> : (
                      <>
                        <PrimaryButton text="Finalizar Obra" iconProps={{ iconName: "Completed" }} onClick={() => handleAccionObra(obraSeleccionada.Id, "finalizar")} className={styles.btnNuevaObra} />
                        <DefaultButton text="Cancelar Obra" iconProps={{ iconName: "Clear" }} onClick={() => handleAccionObra(obraSeleccionada.Id, "cancelar")} />
                        <IconButton iconProps={{ iconName: "Delete" }} title="Eliminar Obra" onClick={() => handleAccionObra(obraSeleccionada.Id, "eliminar")} className={styles.btnClose} />
                      </>
                    )}
                  </Stack>
                </Stack>
              </div>
            </div>
          ) : (
            <div className={styles.emptyState}>
              <Icon iconName="ProjectCollection" className={styles.emptyIcon} />
              <Text variant="xLarge">Selecciona una obra</Text>
              <Text variant="medium">Pincha en un proyecto de la lista para ver su información detallada.</Text>
            </div>
          )}
        </div>
      </div>

      <Modal 
        isOpen={isOpen} 
        onDismiss={() => setIsOpen(false)} 
        containerClassName={styles.modalContainer}
        allowTouchBodyScroll={true} 
        layerProps={{ eventBubblingEnabled: true }} 
      >
        <div className={styles.modalContent}>
          <div className={styles.modalHeader}>
            <Text variant="large" className={styles.modalTitle}>{obraEditandoId ? "🔧 Modificar Parámetros de Obra" : "🚀 Lanzamiento de Nuevo Frente de Obra"}</Text>
            <IconButton iconProps={{ iconName: "Cancel" }} className={styles.btnClose} onClick={() => setIsOpen(false)} />
          </div>
          <Separator className={styles.modalSeparator} />
          <div className={styles.modalBody}>
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField label="Nombre del Proyecto / Frente" required value={nuevaObra.Nombre} onChange={(_, v) => setNuevaObra({ ...nuevaObra, Nombre: v || "" })} />
              
              <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
                <Stack.Item grow={1}>
                  <Dropdown 
                    label="Cliente" 
                    required 
                    options={clientes} 
                    selectedKey={nuevaObra.ClienteId || null} 
                    placeholder="Selecciona un cliente..."
                    onChange={(_, opt) => setNuevaObra({ ...nuevaObra, ClienteId: opt?.key as number })} 
                  />
                </Stack.Item>
                <IconButton 
                  iconProps={{ iconName: "Add" }} 
                  title="Crear nuevo cliente sobre la marcha" 
                  onClick={() => setIsNuevoClienteOpen(true)}
                  styles={{ root: { marginBottom: '2px', backgroundColor: '#f3f2f1', borderRadius: '4px', height: 32, width: 32 } }}
                />
              </Stack>

              <TextField label="Dirección de Obra" value={nuevaObra.Direccion} onChange={(_, v) => setNuevaObra({ ...nuevaObra, Direccion: v || "" })} prefix="📍" />
              
              <div>
                <div
                  ref={mapRef}
                  data-is-scrollable="true" 
                  style={{ width: "100%", height: "150px", borderRadius: "8px", border: "1px solid #ddd", overflow: "hidden", marginTop: "4px", position: "relative", zIndex: 1 }}
                />
                <Text variant="xSmall" style={{ color: "#aaa", marginTop: "3px", display: "block" }}>
                  📌 El mapa se centra al escribir la dirección · Clic para mover el marcador
                </Text>
              </div>

              <Stack horizontal tokens={{ childrenGap: 20 }}>
                <DatePicker label="Fecha Inicio" value={nuevaObra.FechaInicio} onSelectDate={(d) => setNuevaObra({ ...nuevaObra, FechaInicio: d || new Date() })} strings={stringsEspanol} firstDayOfWeek={DayOfWeek.Monday} formatDate={(date) => date ? date.toLocaleDateString() : ''} styles={{ root: { flex: 1 } }} />
                <DatePicker label="Fecha Fin Prevista" value={nuevaObra.FechaFin} onSelectDate={(d) => setNuevaObra({ ...nuevaObra, FechaFin: d || new Date() })} strings={stringsEspanol} firstDayOfWeek={DayOfWeek.Monday} formatDate={(date) => date ? date.toLocaleDateString() : ''} styles={{ root: { flex: 1 } }} />
              </Stack>
              {/* CORREGIDO AQUÍ: Guardamos correctamente la propiedad "JornadasTotales" en el onChange en lugar de "Jornadas" */}
              <TextField 
                label="Jornadas Presupuestadas" 
                type="number" 
                required 
                value={nuevaObra.JornadasTotales.toString()} 
                onChange={(_, v) => setNuevaObra({ ...nuevaObra, JornadasTotales: parseInt(v || "0") })} 
              />
              <TextField label="Descripción" multiline rows={3} value={nuevaObra.Descripcion} onChange={(_, v) => setNuevaObra({ ...nuevaObra, Descripcion: v || "" })} />
            </Stack>
          </div>
          <div className={styles.modalFooter}>
            {saving ? <Spinner label="Guardando..." /> : (
              <>
                <DefaultButton text="Cancelar" onClick={() => setIsOpen(false)} />
                <PrimaryButton text={obraEditandoId ? "Actualizar" : "Lanzar Proyecto"} className={styles.btnLaunch} onClick={handleGuardar} disabled={!nuevaObra.Nombre || !nuevaObra.ClienteId} />
              </>
            )}
          </div>
        </div>
      </Modal>

      <Dialog
        hidden={!isNuevoClienteOpen}
        onDismiss={() => setIsOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Añadir Cliente Exprés',
          subText: 'Introduce el nombre de la empresa o cliente para darlo de alta instantáneamente en el sistema.'
        }}
        modalProps={{ isBlocking: true }}
      >
        <TextField 
          label="Nombre del Cliente / Razón Social" 
          required 
          value={nuevoClienteNombre} 
          onChange={(_, v) => setNuevoClienteNombre(v || "")}
          placeholder="Ej: Iberdrola Renovables"
        />
        <DialogFooter>
          {creandoCliente ? (
            <Spinner size={SpinnerSize.medium} label="Registrando..." />
          ) : (
            <>
              <PrimaryButton onClick={handleCrearClienteRapido} text="Crear y Seleccionar" disabled={!nuevoClienteNombre.trim()} />
              <DefaultButton onClick={() => setIsNuevoClienteOpen(false)} text="Cancelar" />
            </>
          )}
        </DialogFooter>
      </Dialog>

      {/* ─── MODAL TRANSPARENTE DE PREVISUALIZACIÓN DE FOTOS (LIGHTBOX) ─── */}
      <Modal
        isOpen={!!fotoAmpliada}
        onDismiss={() => setFotoAmpliada(null)}
        isBlocking={false}
        styles={{
          main: {
            background: 'transparent',
            boxShadow: 'none',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            outline: 'none'
          }
        }}
      >
        {fotoAmpliada && (
          <div style={{ position: 'relative', outline: 'none' }}>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => setFotoAmpliada(null)}
              title="Cerrar imagen"
              styles={{
                root: { position: 'absolute', top: -40, right: -40, color: 'white', zIndex: 10 },
                icon: { fontSize: 20 }
              }}
            />
            <img 
              src={fotoAmpliada} 
              style={{ maxWidth: '90vw', maxHeight: '90vh', borderRadius: '8px', boxShadow: '0 0 20px rgba(0,0,0,0.5)' }} 
              alt="Previsualización ampliada" 
            />
          </div>
        )}
      </Modal>
    </div>
  );
};