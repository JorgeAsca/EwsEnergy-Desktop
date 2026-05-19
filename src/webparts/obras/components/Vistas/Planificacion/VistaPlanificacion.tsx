import * as React from "react";
import {
  Stack,
  Text,
  Persona,
  PersonaSize,
  Spinner,
  SpinnerSize,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  IconButton,
  ProgressIndicator,
  Dropdown,
  IDropdownOption,
  Link,
} from "@fluentui/react";
import { ProjectService } from "../../../service/ProjectService";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { IObra } from "../../../models/IObra";
import { IPersonal } from "../../../models/IPersonal";
import { IAsignacion } from "../../../models/IAsignacion";
import { SPFI } from "@pnp/sp";
import styles from "./VistaPlanificacion.module.scss";

const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

interface IObraPendiente {
  nombre: string;
  motivo: string;
}

interface IVistaPlanificacionProps {
  sp: SPFI;
}

// Función auxiliar para obtener el Lunes de cualquier fecha dada
const obtenerLunes = (d: Date): Date => {
  const date = new Date(d);
  date.setHours(0, 0, 0, 0);
  const day = date.getDay() || 7;
  date.setDate(date.getDate() - day + 1);
  return date;
};

export const VistaPlanificacion: React.FC<IVistaPlanificacionProps> = (props) => {
  const [obras, setObras] = React.useState<IObra[]>([]);
  const [personalDisponible, setPersonalDisponible] = React.useState<IPersonal[]>([]);
  const [asignaciones, setAsignaciones] = React.useState<IAsignacion[]>([]);
  const [loading, setLoading] = React.useState(true);

  // Controla qué semana estamos viendo (siempre apunta al Lunes de esa semana)
  const [fechaInicioSemana, setFechaInicioSemana] = React.useState<Date>(obtenerLunes(new Date()));

  const [selectedAsig, setSelectedAsig] = React.useState<{
    asig: IAsignacion;
    persona: IPersonal;
  } | null>(null);

  const [editPersonId, setEditPersonId] = React.useState<number | null>(null);

  const [obrasPendientes, setObrasPendientes] = React.useState<IObraPendiente[]>([]);
  const [showAddPending, setShowAddPending] = React.useState(false);
  const [newPending, setNewPending] = React.useState<IObraPendiente>({
    nombre: "",
    motivo: "",
  });

  const services = React.useMemo(
    () => ({
      project: new ProjectService(props.sp),
      personal: new PersonalService(props.sp),
      asignaciones: new AsignacionesService(props.sp),
    }),
    [props.sp],
  );

  const cargarDatos = async () => {
    setLoading(true);
    try {
      const [o, p, a] = await Promise.all([
        services.project.getObras(),
        services.personal.getPersonal(),
        services.asignaciones.getAsignaciones(),
      ]);
      setObras(o);
      setPersonalDisponible(p);
      setAsignaciones(a);
    } catch (error) {
      console.error("Error al cargar planificación:", error);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    if (props.sp) {
      cargarDatos();
    }
  }, [props.sp]);

  // Controles de navegación del calendario
  const irSemanaAnterior = () => {
    const nuevaFecha = new Date(fechaInicioSemana);
    nuevaFecha.setDate(nuevaFecha.getDate() - 7);
    setFechaInicioSemana(nuevaFecha);
  };

  const irSemanaSiguiente = () => {
    const nuevaFecha = new Date(fechaInicioSemana);
    nuevaFecha.setDate(nuevaFecha.getDate() + 7);
    setFechaInicioSemana(nuevaFecha);
  };

  const irHoy = () => {
    setFechaInicioSemana(obtenerLunes(new Date()));
  };

  // Calculamos dinámicamente los 5 días de la semana actual que estamos viendo
  const diasDeLaSemanaActual = React.useMemo(() => {
    return [0, 1, 2, 3, 4].map(offset => {
      const d = new Date(fechaInicioSemana);
      d.setDate(d.getDate() + offset);
      return {
        nombre: DIAS_SEMANA[offset],
        fechaObj: d,
        diaNumero: d.getDate(),
        mesTexto: d.toLocaleString('es-ES', { month: 'short' })
      };
    });
  }, [fechaInicioSemana]);

  const onDrop = async (ev: React.DragEvent, obraId: number, fechaExacta: Date) => {
    ev.preventDefault();
    const personId = parseInt(ev.dataTransfer.getData("personId"));
    try {
      await services.asignaciones.asignarPersonal({
        ObraId: obraId,
        PersonalId: personId,
        FechaInicio: fechaExacta,
        FechaFinPrevista: fechaExacta,
        EstadoProgreso: 0,
      });
      await cargarDatos();
    } catch (error) {
      console.error("Error al asignar personal:", error);
    }
  };

  const editarAsignacion = async () => {
    if (!selectedAsig?.asig.Id || !editPersonId) return;
    try {
      await services.asignaciones.actualizarAsignacion(selectedAsig.asig.Id, {
        PersonalId: editPersonId
      });
      setSelectedAsig(null);
      await cargarDatos();
    } catch (error) {
      console.error("Error al editar asignación:", error);
    }
  };

  const eliminarAsignacion = async () => {
    if (!selectedAsig?.asig.Id) return;
    try {
      await services.asignaciones.eliminarAsignacion(selectedAsig.asig.Id);
      setSelectedAsig(null);
      await cargarDatos();
    } catch (error) {
      console.error("Error al eliminar asignación:", error);
    }
  };

  const personalOptions: IDropdownOption[] = personalDisponible.map(p => ({
    key: p.Id,
    text: p.NombreyApellido
  }));

  const mesAnioActualTexto = fechaInicioSemana.toLocaleString('es-ES', { month: 'long', year: 'numeric' }).toUpperCase();

  if (loading) return <Spinner label="Cargando planificación semanal..." size={SpinnerSize.large} />;

  return (
    <Stack tokens={{ childrenGap: 15 }} className={styles.vistaPlanificacion}>
      {/* HEADER PRINCIPAL */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" className={styles.titulo}>
          Planificación Semanal 📅
        </Text>
        <PrimaryButton
          iconProps={{ iconName: "Add" }}
          text="Nota Pendiente"
          onClick={() => setShowAddPending(true)}
        />
      </Stack>

      {/* PANEL PERSONAL ARRIBA */}
      <div className={styles.personalPanelTop}>
        <div className={styles.personalListHorizontal}>
          {personalDisponible.map((p) => (
            <div
              key={p.Id}
              draggable
              onDragStart={(e) => e.dataTransfer.setData("personId", p.Id.toString())}
              className={styles.draggablePersonaCard}
            >
              <Persona text={p.NombreyApellido} imageUrl={p.FotoPerfil} size={PersonaSize.size24} />
            </div>
          ))}
        </div>
      </div>

      {/* CONTROLES DEL CALENDARIO */}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} style={{ padding: '10px 0', backgroundColor: '#fff', borderRadius: '8px', paddingLeft: '15px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
        <IconButton iconProps={{ iconName: "ChevronLeft" }} title="Semana anterior" onClick={irSemanaAnterior} />
        <DefaultButton text="Hoy" onClick={irHoy} />
        <IconButton iconProps={{ iconName: "ChevronRight" }} title="Semana siguiente" onClick={irSemanaSiguiente} />
        <Text variant="large" style={{ fontWeight: 600, marginLeft: '15px', color: '#0078d4' }}>
          {mesAnioActualTexto}
        </Text>
      </Stack>

      {/* CUERPO: TABLA Y PENDIENTES */}
      <Stack horizontal tokens={{ childrenGap: 15 }} styles={{ root: { width: "100%", alignItems: "start" } }}>
        <div className={styles.tableContainer}>
          <table className={styles.planTable}>
            <thead>
              <tr>
                <th className={styles.colObra}>Obra</th>
                {diasDeLaSemanaActual.map((dia) => (
                  <th key={dia.nombre} className={styles.colDia} style={{ textAlign: 'center', padding: '10px 0' }}>
                    <div style={{ fontWeight: 600 }}>{dia.nombre}</div>
                    <div style={{ fontSize: '13px', fontWeight: 'normal', color: '#605e5c', marginTop: '2px' }}>
                      {dia.diaNumero} {dia.mesTexto}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {obras.map((obra) => (
                <tr key={obra.Id}>
                  <td className={styles.cellObra}>
                    <Stack tokens={{ childrenGap: 4 }}>
                      
                      {/* ENLACE CLICKEABLE A LA OBRA EN LA MISMA PESTAÑA */}
                      <Link 
                        styles={{ root: { fontWeight: 600, fontSize: '14px', color: '#0078d4', textDecoration: 'none' } }}
                        onClick={() => {
                          const urlSharePoint = `${window.location.origin}/sites/EWSStockManagement/Lists/Obras/DispForm.aspx?ID=${obra.Id}`;
                          window.location.href = urlSharePoint;
                        }}
                      >
                        {obra.Title}
                      </Link>

                      <Stack>
                        <Text variant="small" styles={{ root: { color: '#666', fontSize: '11px' } }}>
                          Avance: {obra.ProgresoReal || 0}% • {obra.EstadoObra}
                        </Text>
                        <ProgressIndicator 
                          percentComplete={(obra.ProgresoReal || 0) / 100} 
                          styles={{ itemProgress: { padding: 0 }, progressBar: { backgroundColor: '#107c41' } }} 
                        />
                      </Stack>
                    </Stack>
                  </td>

                  {/* Celdas de los días de la semana dinámica */}
                  {diasDeLaSemanaActual.map((dia) => {
                    const asigsEnDia = asignaciones.filter((a) => {
                      if (a.ObraId !== obra.Id || !a.FechaInicio) return false;
                      const fechaAsig = new Date(a.FechaInicio);
                      return (
                        fechaAsig.getFullYear() === dia.fechaObj.getFullYear() &&
                        fechaAsig.getMonth() === dia.fechaObj.getMonth() &&
                        fechaAsig.getDate() === dia.fechaObj.getDate()
                      );
                    });

                    return (
                      <td
                        key={dia.nombre}
                        onDragOver={(e) => e.preventDefault()}
                        onDrop={(e) => onDrop(e, obra.Id, dia.fechaObj)}
                        className={styles.dropZone}
                      >
                        <div className={styles.asignadosConsola}>
                          {asigsEnDia.map((a) => {
                            const p = personalDisponible.find((pers) => pers.Id === a.PersonalId);
                            return p ? (
                              <div
                                key={a.Id}
                                onClick={() => {
                                  setSelectedAsig({ asig: a, persona: p });
                                  setEditPersonId(p.Id); 
                                }}
                                className={styles.fotoAsignada}
                              >
                                <Persona text={p.NombreyApellido} imageUrl={p.FotoPerfil} size={PersonaSize.size32} />
                              </div>
                            ) : null;
                          })}
                        </div>
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* PANEL PENDIENTES */}
        <div className={styles.pendingPanel}>
          <Text className={styles.panelTituloCompacto}>Pendientes</Text>
          <div className={styles.pendingList}>
            {obrasPendientes.length === 0 && <span className={styles.emptyText}>Sin notas</span>}
            {obrasPendientes.map((op, idx) => (
              <div key={idx} className={styles.pendingItem}>
                <Stack horizontal horizontalAlign="space-between">
                  <Text className={styles.pendingName}>{op.nombre}</Text>
                  <IconButton
                    iconProps={{ iconName: "Cancel" }}
                    styles={{ root: { height: 16, width: 16, fontSize: 10 } }}
                    onClick={() => setObrasPendientes(obrasPendientes.filter((_, i) => i !== idx))}
                  />
                </Stack>
                <Text className={styles.pendingReason}>{op.motivo}</Text>
              </div>
            ))}
          </div>
        </div>
      </Stack>

      {/* DIALOG AÑADIR PENDIENTE */}
      <Dialog hidden={!showAddPending} onDismiss={() => setShowAddPending(false)} dialogContentProps={{ type: DialogType.normal, title: "Nueva Nota Pendiente" }}>
        <TextField label="Nombre" value={newPending.nombre} onChange={(_, v) => setNewPending({ ...newPending, nombre: v || "" })} />
        <TextField label="Motivo" multiline rows={3} value={newPending.motivo} onChange={(_, v) => setNewPending({ ...newPending, motivo: v || "" })} />
        <DialogFooter>
          <PrimaryButton onClick={() => { setObrasPendientes([...obrasPendientes, newPending]); setNewPending({ nombre: "", motivo: "" }); setShowAddPending(false); }} text="Añadir" />
          <DefaultButton onClick={() => setShowAddPending(false)} text="Cancelar" />
        </DialogFooter>
      </Dialog>

      {/* DIALOG EDITAR / ELIMINAR ASIGNACIÓN */}
      <Dialog
        hidden={!selectedAsig}
        onDismiss={() => setSelectedAsig(null)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Editar Asignación",
          subText: "Cambia el trabajador asignado o elimina la asignación por completo."
        }}
      >
        <Stack tokens={{ childrenGap: 15 }} style={{ marginTop: '10px' }}>
            <Dropdown
              label="Trabajador asignado"
              selectedKey={editPersonId}
              options={personalOptions}
              onChange={(_, option) => setEditPersonId(option?.key as number)}
            />
        </Stack>
        <DialogFooter>
          <PrimaryButton 
            onClick={editarAsignacion} 
            text="Guardar Cambios" 
            disabled={editPersonId === selectedAsig?.persona.Id} 
          />
          <DefaultButton 
            onClick={eliminarAsignacion} 
            text="Eliminar" 
            styles={{ root: { color: '#d13438' } }} 
          />
          <DefaultButton onClick={() => setSelectedAsig(null)} text="Cancelar" />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};