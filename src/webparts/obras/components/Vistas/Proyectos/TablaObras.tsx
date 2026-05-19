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
} from "@fluentui/react";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { MapContainer, TileLayer, Marker, useMapEvents, useMap } from 'react-leaflet';
import L from 'leaflet';
import { SPComponentLoader } from '@microsoft/sp-loader';

import { ProjectService } from "../../../service/ProjectService";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { IObraCard } from "../../../models/IObraCard";
import styles from "./TablaObras.module.scss";

// 1. Cargamos el CSS de Leaflet desde un CDN seguro para evitar errores de Webpack en SPFx
SPComponentLoader.loadCss('https://unpkg.com/leaflet@1.9.4/dist/leaflet.css');

// 2. Apuntamos las imágenes de los marcadores al CDN
let DefaultIcon = L.icon({
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
  iconSize: [25, 41],
  iconAnchor: [12, 41]
});
L.Marker.prototype.options.icon = DefaultIcon;

// Componente auxiliar para manejar clics y geocodificación inversa
const LocationPicker: React.FC<{ 
    position: [number, number], 
    setPosition: (pos: [number, number]) => void,
    setAddress: (addr: string) => void 
}> = ({ position, setPosition, setAddress }) => {
    const map = useMap();

    useMapEvents({
        // Tipamos 'e' como L.LeafletMouseEvent
        click: async (e: L.LeafletMouseEvent) => {
            const { lat, lng } = e.latlng;
            setPosition([lat, lng]);
            map.flyTo(e.latlng, map.getZoom());

            // Geocodificación inversa gratuita con Nominatim (OpenStreetMap)
            try {
                const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}`);
                const data = await response.json();
                if (data.display_name) {
                    setAddress(data.display_name);
                }
            } catch (error) {
                console.error("Error obteniendo dirección:", error);
            }
        },
    });

    return <Marker position={position} />;
};

// Componente para centrar el mapa cuando cambia la posición externamente
const ChangeView = ({ center }: { center: [number, number] }) => {
    const map = useMap();
    map.setView(center);
    return null;
};

export interface ITablaObrasProps {
  sp: SPFI;
}

export const TablaObras: React.FC<ITablaObrasProps> = (props) => {
  // --- ESTADOS DE DATOS ---
  const [obras, setObras] = React.useState<IObraCard[]>([]);
  const [clientes, setClientes] = React.useState<IDropdownOption[]>([]);
  const [obraSeleccionada, setObraSeleccionada] = React.useState<IObraCard | null>(null);
  const [fotosObra, setFotosObra] = React.useState<any[]>([]);

  // --- ESTADOS DE CONTROL ---
  const [loading, setLoading] = React.useState(true);
  const [loadingFotos, setLoadingFotos] = React.useState(false);
  const [isOpen, setIsOpen] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [obraEditandoId, setObraEditandoId] = React.useState<number | null>(null);

  // --- ESTADOS DE MAPA ---
  const [formCoords, setFormCoords] = React.useState<[number, number]>([40.4167, -3.7037]); // Coordenadas para Modal
  const [detailCoords, setDetailCoords] = React.useState<[number, number]>([40.4167, -3.7037]); // Coordenadas para Detalle

  // --- ESTADO DE FORMULARIO ---
  const [nuevaObra, setNuevaObra] = React.useState({
    Nombre: "",
    Descripcion: "",
    ClienteId: 0,
    Direccion: "",
    EstadoObra: "Pendiente",
    FechaInicio: new Date(),
    FechaFin: new Date(),
    JornadasTotales: 30,
  });

  const estadoOptions: IDropdownOption[] = [
    { key: "Pendiente", text: "Pendiente" },
    { key: "En Proceso", text: "En Proceso" },
    { key: "Completado", text: "Completado" },
  ];

  // --- SERVICIOS MEMOIZADOS ---
  const services = React.useMemo(() => ({
    project: new ProjectService(props.sp),
    personal: new PersonalService(props.sp),
    asig: new AsignacionesService(props.sp)
  }), [props.sp]);

  // --- LÓGICA DE CARGA CENTRALIZADA ---
  const cargarTodo = async () => {
    try {
      setLoading(true);
      const [listaObras, listaClientes, listaAsignaciones, listaPersonal] = await Promise.all([
        services.project.getObras(),
        props.sp.web.lists.getByTitle('Clientes').items.select("Id", "Title")(), // Petición migrada a PnPjs
        services.asig.getAsignaciones(),
        services.personal.getPersonal(),
      ]);

      const opcionesClientes: IDropdownOption[] = (listaClientes || []).map((c: any) => ({ 
        key: c.Id, 
        text: c.Title 
      }));
      setClientes(opcionesClientes);

      const obrasProcesadas: IObraCard[] = listaObras.map((o) => {
        const porcentajeReal = (o.ProgresoReal || 0) / 100;
        const asigsObra = (listaAsignaciones as any[]).filter(a => Number(a.ObraId) === Number(o.Id));

        const operariosAsignados = Array.from(new Set(asigsObra.map(a => Number(a.PersonalId))))
          .map(pid => {
            const pers = (listaPersonal as any[]).find(p => Number(p.Id) === pid);
            return {
              personaName: pers?.NombreyApellido || "Operario",
              imageUrl: pers?.FotoPerfil || "",
            };
          });

        return {
          ...o,
          clienteNombre: opcionesClientes.find(c => Number(c.key) === (o as any).Cliente?.Id)?.text || "Cliente no definido",
          porcentajeReal: Math.min(Math.max(porcentajeReal, 0), 1),
          operarios: operariosAsignados,
          jornadasConsumidas: parseFloat((porcentajeReal * (o.JornadasTotales || 30)).toFixed(1)),
        } as IObraCard;
      });

      setObras(obrasProcesadas);

      if (obraSeleccionada) {
        const actualizada = obrasProcesadas.find(o => o.Id === obraSeleccionada.Id);
        if (actualizada) setObraSeleccionada(actualizada);
      }
    } catch (e) {
      console.error("Error en Dashboard:", e);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => { 
    if(props.sp) cargarTodo(); 
  }, [props.sp]);

  // --- MAPA: BUSCAR DIRECCIÓN (GEOCODING DIRECTO) ---
  const buscarDireccion = async (addr: string, isDetail: boolean) => {
    if (addr.length < 4) return;
    try {
        const response = await fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(addr)}&limit=1`);
        const data = await response.json();
        if (data && data.length > 0) {
            const pos: [number, number] = [parseFloat(data[0].lat), parseFloat(data[0].lon)];
            isDetail ? setDetailCoords(pos) : setFormCoords(pos);
        }
    } catch (e) { console.error(e); }
  };

  // --- MANEJADORES DE ACCIONES ---
  const verDetallesObra = async (obra: IObraCard) => {
    setObraSeleccionada(obra);
    setLoadingFotos(true);
    
    if (obra.DireccionObra) buscarDireccion(obra.DireccionObra, true);

    try {
      const fotos = await services.project.getFotosPorObra(obra.Id as number);
      setFotosObra(fotos || []);
    } catch (e) {
      console.error(e);
    } finally {
      setLoadingFotos(false);
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
      alert("Error al guardar los cambios.");
    } finally {
      setSaving(false);
    }
  };

  const resetForm = () => {
    setObraEditandoId(null);
    setNuevaObra({
      Nombre: "", Descripcion: "", ClienteId: 0, Direccion: "", EstadoObra: "Pendiente",
      FechaInicio: new Date(), FechaFin: new Date(), JornadasTotales: 30,
    });
    setFormCoords([40.4167, -3.7037]); // Volver al centro por defecto
  };

  const handleAccionObra = async (id: number, accion: 'finalizar' | 'cancelar' | 'eliminar') => {
    const confirmacion = {
      finalizar: "¿Estás seguro de finalizar esta obra?",
      cancelar: "¿Deseas cancelar esta obra? No aparecerá activa.",
      eliminar: "⚠️ ¿ESTÁS SEGURO? Se borrarán todos los registros permanentemente."
    };

    if (!window.confirm(confirmacion[accion])) return;

    try {
      setIsProcessing(true);
      if (accion === 'finalizar') await services.project.finalizarObra(id);
      if (accion === 'cancelar') await services.project.cancelarObra(id);
      if (accion === 'eliminar') {
        // Petición de borrado migrada a PnPjs
        await props.sp.web.lists.getByTitle('Proyectos y Obras').items.getById(id).delete();
      }
      setObraSeleccionada(null);
      await cargarTodo();
    } catch (e) {
      console.error(`Error en ${accion}:`, e);
    } finally {
      setIsProcessing(false);
    }
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

  const obrasAgrupadas = obras.reduce((acc, obra) => {
    const estado = obra.EstadoObra || "Sin Asignar";
    if (!acc[estado]) acc[estado] = [];
    acc[estado].push(obra);
    return acc;
  }, {} as Record<string, IObraCard[]>);

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
        {/* COLUMNA IZQUIERDA: LISTADO AGRUPADO */}
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

        {/* COLUMNA DERECHA: DETALLES */}
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
                <DefaultButton iconProps={{ iconName: "Edit" }} text="Editar" onClick={() => {
                  setObraEditandoId(obraSeleccionada.Id as number);
                  setNuevaObra({
                    Nombre: obraSeleccionada.Title, Descripcion: obraSeleccionada.Descripcion || "",
                    ClienteId: (clientes.find(c => c.text === obraSeleccionada.clienteNombre)?.key as number) || 0,
                    Direccion: obraSeleccionada.DireccionObra || "",
                    EstadoObra: obraSeleccionada.EstadoObra || "Pendiente",
                    FechaInicio: new Date(obraSeleccionada.FechaInicio || Date.now()),
                    FechaFin: new Date(obraSeleccionada.FechaFinPrevista || Date.now()),
                    JornadasTotales: obraSeleccionada.JornadasTotales || 30
                  });
                  if(obraSeleccionada.DireccionObra) buscarDireccion(obraSeleccionada.DireccionObra, false);
                  setIsOpen(true);
                }} />
              </Stack>

              <Separator />

              <Stack horizontal tokens={{ childrenGap: 40 }} className={styles.infoSection}>
                <Stack>
                  <Text className={styles.labelSeccion}>Dirección</Text>
                  <Text><Icon iconName="MapPin" className={styles.iconVerde} /> {obraSeleccionada.DireccionObra || "Sin dirección"}</Text>
                </Stack>
                <Stack>
                  <Text className={styles.labelSeccion}>Jornadas Consumidas</Text>
                  <Text><Icon iconName="Calendar" className={styles.iconVerde} /> {obraSeleccionada.jornadasConsumidas} / {obraSeleccionada.JornadasTotales || 30}</Text>
                </Stack>
                <Stack>
                  <Text className={styles.labelSeccion}>Avance Físico</Text>
                  <Text><Icon iconName="CompletedSolid" className={styles.iconVerde} /> {(obraSeleccionada.porcentajeReal * 100).toFixed(0)}% Ejecutado</Text>
                </Stack>
                <Stack>
                  <Text className={styles.labelSeccion}>Equipo en Campo</Text>
                  {obraSeleccionada.operarios?.length > 0 ? (
                    <Facepile personas={obraSeleccionada.operarios} personaSize={PersonaSize.size32} />
                  ) : <Text variant="small" style={{ fontStyle: "italic", color: "#888" }}>Sin personal asignado</Text>}
                </Stack>
              </Stack>

              {/* MAPA DE VISTA DETALLE */}
              {obraSeleccionada.DireccionObra && (
                 <div style={{ width: '100%', height: '200px', borderRadius: '8px', overflow: 'hidden', border: '1px solid #e1dfdd', marginTop: '15px' }}>
                     <MapContainer center={detailCoords} zoom={15} style={{ height: '100%', width: '100%', zIndex: 0 }}>
                         <ChangeView center={detailCoords} />
                         <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                         <Marker position={detailCoords} />
                     </MapContainer>
                 </div>
              )}

              <div className={styles.planosSection}>
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 15 } }}>
                  <Text variant="large" className={styles.sectionTitle}>Planos y Documentación</Text>
                  <DefaultButton iconProps={{ iconName: "Upload" }} className={styles.btnUpload}>Añadir Archivo</DefaultButton>
                </Stack>
                <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
                  <div className={styles.planoCard}><Icon iconName="PDF" className={styles.pdfIcon} /><Text variant="smallPlus">Esquema_Eléctrico_v2.pdf</Text></div>
                  <div className={styles.planoCard}><Icon iconName="VisioDocument" className={styles.dwgIcon} /><Text variant="smallPlus">Topografía_Terreno.dwg</Text></div>
                </Stack>
              </div>

              <div className={styles.historialSection}>
                <Text variant="large" className={styles.sectionTitle}>Reportes de Jornada</Text>
                {loadingFotos ? <Spinner size={SpinnerSize.large} label="Cargando reportes..." /> :
                  fotosObra.length > 0 ? (
                    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { marginTop: 15 } }}>
                      {fotosObra.map((f, i) => (
                        <div key={i} className={styles.fotoCard}>
                          <Stack horizontal tokens={{ childrenGap: 15 }}>
                            <Image src={f.UrlFoto?.Url} width={120} height={90} imageFit={ImageFit.cover} className={styles.fotoThumb} />
                            <Stack>
                              <Text className={styles.fotoFecha}>📅 {new Date(f.FechaRegistro).toLocaleDateString()} - Worker {f.Operario}</Text>
                              <div className={styles.fotoComentarioBox}><Text className={styles.fotoComentarioText}>"{f.Comentarios || "Sin observaciones técnicas"}"</Text></div>
                            </Stack>
                          </Stack>
                        </div>
                      ))}
                    </Stack>
                  ) : <MessageBar messageBarType={MessageBarType.info}>No hay reportes para esta obra.</MessageBar>}
              </div>

              <div className={styles.planosSection}>
                <Separator />
                <Stack tokens={{ childrenGap: 15 }} style={{ marginTop: '20px' }}>
                  <Text variant="large" className={styles.sectionTitle}>Gestión de Obra</Text>
                  <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
                    {isProcessing ? <Spinner label="Procesando..." /> : (
                      <>
                        <PrimaryButton
                          text="Finalizar Obra"
                          iconProps={{ iconName: 'Completed' }}
                          onClick={() => handleAccionObra(obraSeleccionada.Id, 'finalizar')}  
                          className={styles.btnNuevaObra}
                        />
                        <DefaultButton
                          text="Cancelar Obra"
                          iconProps={{ iconName: 'Clear' }}
                          onClick={() => handleAccionObra(obraSeleccionada.Id, 'cancelar')}
                        />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Eliminar Obra" onClick={() => handleAccionObra(obraSeleccionada.Id, 'eliminar')} className={styles.btnClose} />
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

      {/* MODAL DE CREACIÓN / EDICIÓN */}
      <Modal isOpen={isOpen} onDismiss={() => setIsOpen(false)} containerClassName={styles.modalFlotanteContainer}>
        <div className={styles.modalContent}>
          <div className={styles.modalHeader}>
            <Text variant="xLarge" className={styles.modalTitle}>{obraEditandoId ? "Editar Proyecto" : "Configurar Nuevo Proyecto"}</Text>
            <IconButton iconProps={{ iconName: "Cancel" }} onClick={() => setIsOpen(false)} className={styles.btnClose} />
          </div>
          <Separator className={styles.modalSeparator} />
          <div className={styles.modalBody}>
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField label="Nombre del Proyecto" required value={nuevaObra.Nombre} onChange={(_, v) => setNuevaObra({ ...nuevaObra, Nombre: v || "" })} />
              <Dropdown label="Cliente" required options={clientes} selectedKey={nuevaObra.ClienteId} onChange={(_, opt) => setNuevaObra({ ...nuevaObra, ClienteId: opt?.key as number })} />
              
              <TextField 
                label="Dirección / Ubicación" 
                value={nuevaObra.Direccion} 
                onChange={(_, v) => {
                  setNuevaObra({ ...nuevaObra, Direccion: v || "" });
                  buscarDireccion(v || "", false);
                }} 
                placeholder="Escribe la dirección o selecciona en el mapa"
              />

              {/* MAPA INTERACTIVO DE FORMULARIO */}
              <div style={{ width: '100%', height: '220px', borderRadius: '8px', overflow: 'hidden', border: '1px solid #e1dfdd' }}>
                  <MapContainer center={formCoords} zoom={13} style={{ height: '100%', width: '100%', zIndex: 0 }}>
                      <ChangeView center={formCoords} />
                      <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" />
                      <LocationPicker 
                        position={formCoords} 
                        setPosition={setFormCoords} 
                        setAddress={(addr) => setNuevaObra({ ...nuevaObra, Direccion: addr })} 
                      />
                  </MapContainer>
              </div>

              <Dropdown label="Estado" required options={estadoOptions} selectedKey={nuevaObra.EstadoObra} onChange={(_, opt) => setNuevaObra({ ...nuevaObra, EstadoObra: opt?.key as string })} />
              
              <Stack horizontal tokens={{ childrenGap: 20 }}>
                <TextField label="Jornadas Presupuestadas" type="number" required value={nuevaObra.JornadasTotales.toString()} onChange={(_, v) => setNuevaObra({ ...nuevaObra, JornadasTotales: parseInt(v || "0") })} styles={{ root: { flex: 1 } }} />
                <DatePicker label="Fecha Inicio" value={nuevaObra.FechaInicio} onSelectDate={(d) => setNuevaObra({ ...nuevaObra, FechaInicio: d || new Date() })} styles={{ root: { flex: 1 } }} />
              </Stack>
            </Stack>
          </div>
          <div className={styles.modalFooter}>
            <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="end">
              {saving ? <Spinner label="Guardando..." /> : (
                <>
                  <PrimaryButton text={obraEditandoId ? "Actualizar" : "Lanzar Proyecto"} onClick={handleGuardar} disabled={!nuevaObra.Nombre || !nuevaObra.ClienteId} />
                  <DefaultButton text="Cancelar" onClick={() => setIsOpen(false)} />
                </>
              )}
            </Stack>
          </div>
        </div>
      </Modal>
    </div>
  );
};