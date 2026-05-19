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
import { SPHttpClient } from "@microsoft/sp-http";
import { ProjectService } from "../../../service/ProjectService";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { IObraCard } from "../../../models/IObraCard";
import styles from "./TablaObras.module.scss";

export const TablaObras: React.FC<{ context: any }> = (props) => {
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

  // --- ESTADO DE FORMULARIO ---
  const [nuevaObra, setNuevaObra] = React.useState({
    Nombre: "",
    Descripcion: "",
    ClienteId: 0,
    Direccion: "",
    FechaInicio: new Date(),
    FechaFin: new Date(),
    JornadasTotales: 30,
  });

  // --- SERVICIOS MEMOIZADOS ---
  const services = React.useMemo(() => ({
    project: new ProjectService(props.context),
    personal: new PersonalService(props.context),
    asig: new AsignacionesService(props.context)
  }), [props.context]);

  // --- LÓGICA DE CARGA CENTRALIZADA ---
  const cargarTodo = async () => {
    try {
      setLoading(true);
      const [listaObras, respClientes, listaAsignaciones, listaPersonal] = await Promise.all([
        services.project.getObras(),
        props.context.spHttpClient.get(
          `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Clientes')/items?$select=Id,Title`,
          SPHttpClient.configurations.v1
        ),
        services.asig.getAsignaciones(),
        services.personal.getPersonal(),
      ]);

      let opcionesClientes: IDropdownOption[] = [];
      if (respClientes.ok) {
        const dataC = await respClientes.json();
        opcionesClientes = (dataC.value || []).map((c: any) => ({ key: c.Id, text: c.Title }));
        setClientes(opcionesClientes);
      }

      // Transformación de datos (Preservando tu lógica de cálculo de jornadas y equipo)
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
        };
      });

      setObras(obrasProcesadas);

      // Actualizar selección actual si existe
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

  React.useEffect(() => { cargarTodo(); }, []);

  // --- MANEJADORES DE ACCIONES ---
  const verDetallesObra = async (obra: IObraCard) => {
    setObraSeleccionada(obra);
    setLoadingFotos(true);
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
      Nombre: "", Descripcion: "", ClienteId: 0, Direccion: "",
      FechaInicio: new Date(), FechaFin: new Date(), JornadasTotales: 30,
    });
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
        const endpoint = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Proyectos y Obras')/items(${id})`;
        await props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json', 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' }
        });
      }
      setObraSeleccionada(null);
      await cargarTodo();
    } catch (e) {
      console.error(`Error en ${accion}:`, e);
    } finally {
      setIsProcessing(false);
    }
  };

  // --- RENDERIZADO DE APOYO ---
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
        {/* COLUMNA IZQUIERDA: LISTADO */}
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

        {/* COLUMNA DERECHA: DETALLE */}
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
                    FechaInicio: new Date(obraSeleccionada.FechaInicio || Date.now()),
                    FechaFin: new Date(obraSeleccionada.FechaFinPrevista || Date.now()),
                    JornadasTotales: obraSeleccionada.JornadasTotales || 30
                  });
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
              <TextField label="Dirección de Obra" value={nuevaObra.Direccion} onChange={(_, v) => setNuevaObra({ ...nuevaObra, Direccion: v || "" })} />
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