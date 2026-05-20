import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
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
  Icon,
} from "@fluentui/react";
import { ProjectService } from "../../../service/ProjectService";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { IObra } from "../../../models/IObra";
import { IPersonal } from "../../../models/IPersonal";
import { IAsignacion } from "../../../models/IAsignacion";
import styles from "./VistaPlanificacion.module.scss";

const DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"];

interface IObraPendiente {
  nombre: string;
  motivo: string;
}

export const VistaPlanificacion: React.FC<{ context: WebPartContext }> = ({ context }) => {
  const [obras, setObras] = React.useState<IObra[]>([]);
  const [personalDisponible, setPersonalDisponible] = React.useState<IPersonal[]>([]);
  const [asignaciones, setAsignaciones] = React.useState<IAsignacion[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [selectedAsig, setSelectedAsig] = React.useState<{ asig: IAsignacion; persona: IPersonal } | null>(null);
  const [obrasPendientes, setObrasPendientes] = React.useState<IObraPendiente[]>([]);
  const [showAddPending, setShowAddPending] = React.useState(false);
  const [newPending, setNewPending] = React.useState<IObraPendiente>({ nombre: "", motivo: "" });

  // ESTADO PARA NAVEGACIÓN DE SEMANAS
  const [fechaBase, setFechaBase] = React.useState(new Date());

  const services = React.useMemo(() => ({
    project: new ProjectService(context),
    personal: new PersonalService(context),
    asig: new AsignacionesService(context),
  }), [context]);

  // Lógica para obtener el lunes de la semana basado en fechaBase
  const obtenerLunes = (d: Date) => {
    const fecha = new Date(d);
    const dia = fecha.getDay();
    const dif = fecha.getDate() - dia + (dia === 0 ? -6 : 1);
    fecha.setDate(dif);
    return fecha;
  };

  const lunesSemana = obtenerLunes(fechaBase);

  // Obtener fecha específica para una columna de día
  const getFechaPorDiaIndex = (index: number): Date => {
    const fecha = new Date(lunesSemana);
    fecha.setDate(lunesSemana.getDate() + index);
    return fecha;
  };

  const cargarDatos = async () => {
    setLoading(true);
    try {
      const [o, p, a] = await Promise.all([
        services.project.getObras(),
        services.personal.getPersonal(),
        services.asig.getAsignaciones(),
      ]);
      setObras(o);
      setPersonalDisponible(p);
      setAsignaciones(a);
    } catch (error) {
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => { cargarDatos(); }, []);

  // --- FUNCIÓN ONDROP MODIFICADA ---
  const onDrop = async (ev: React.DragEvent, obraId: number, indexDia: number) => {
    ev.preventDefault();
    
    // Obtenemos qué tipo de elemento estamos arrastrando
    const personIdStr = ev.dataTransfer.getData("personId"); // Viene del panel superior (nueva asignación)
    const assignmentIdStr = ev.dataTransfer.getData("assignmentId"); // Viene de la tabla (mover asignación)
    
    const fecha = getFechaPorDiaIndex(indexDia);
    const fechaFormateada = fecha.toDateString();

    // 1. SI ESTAMOS MOVIENDO UNA ASIGNACIÓN EXISTENTE
    if (assignmentIdStr) {
      const assignmentId = parseInt(assignmentIdStr, 10);
      const asigToMove = asignaciones.find(a => a.Id === assignmentId);
      
      if (asigToMove) {
        // Evitar hacer nada si lo sueltas exactamente en la misma celda de la que viene
        if (asigToMove.ObraId === obraId && new Date(asigToMove.FechaInicio).toDateString() === fechaFormateada) {
          return; 
        }

        // Comprobar que no esté ya duplicado en el destino
        const isDuplicate = asignaciones.some(a => 
          a.ObraId === obraId && 
          new Date(a.FechaInicio).toDateString() === fechaFormateada && 
          a.PersonalId === asigToMove.PersonalId &&
          a.Id !== assignmentId // Excluimos la que estamos moviendo
        );
        
        if (isDuplicate) {
          alert("Este trabajador ya está asignado a esta obra ese mismo día.");
          return;
        }

        setLoading(true);
        // Simulamos un movimiento: Eliminamos la antigua y creamos la nueva en la nueva celda
        await services.asig.eliminarAsignacion(assignmentId);
        await services.asig.asignarPersonal({
          ObraId: obraId,
          PersonalId: asigToMove.PersonalId,
          FechaInicio: fecha,
          FechaFinPrevista: fecha,
          EstadoProgreso: 0,
        });
        await cargarDatos();
      }
      return;
    }

    // 2. SI ESTAMOS CREANDO UNA NUEVA ASIGNACIÓN DESDE ARRIBA
    if (personIdStr) {
      const personId = parseInt(personIdStr, 10);

      // Comprobar duplicados
      const isDuplicate = asignaciones.some(a => 
        a.ObraId === obraId && 
        new Date(a.FechaInicio).toDateString() === fechaFormateada && 
        a.PersonalId === personId
      );

      if (isDuplicate) {
        alert("Este trabajador ya está asignado a esta obra ese mismo día.");
        return;
      }

      setLoading(true);
      await services.asig.asignarPersonal({
        ObraId: obraId,
        PersonalId: personId,
        FechaInicio: fecha,
        FechaFinPrevista: fecha,
        EstadoProgreso: 0,
      });
      await cargarDatos();
    }
  };

  const eliminarAsignacion = async () => {
    if (!selectedAsig?.asig.Id) return;
    setLoading(true);
    await services.asig.eliminarAsignacion(selectedAsig.asig.Id);
    setSelectedAsig(null);
    await cargarDatos();
  };

  if (loading) return <Spinner label="Cargando planificación..." size={SpinnerSize.large} />;

  return (
    <Stack tokens={{ childrenGap: 15 }} className={styles.vistaPlanificacion}>
      
      {/* HEADER CON NAVEGACIÓN */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" className={styles.titulo}>Planificación Semanal</Text>
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <DefaultButton iconProps={{ iconName: 'ChevronLeft' }} onClick={() => setFechaBase(prev => new Date(prev.setDate(prev.getDate() - 7)))} />
          <Text variant="medium" style={{ alignSelf: 'center', fontWeight: 'bold' }}>
            {lunesSemana.toLocaleDateString()} al {getFechaPorDiaIndex(4).toLocaleDateString()}
          </Text>
          <DefaultButton iconProps={{ iconName: 'ChevronRight' }} onClick={() => setFechaBase(prev => new Date(prev.setDate(prev.getDate() + 7)))} />
          <PrimaryButton iconProps={{ iconName: "Add" }} text="Nota Pendiente" onClick={() => setShowAddPending(true)} />
        </Stack>
      </Stack>

      <div className={styles.personalPanelTop}>
        <div className={styles.personalListHorizontal}>
          {personalDisponible.map((p) => (
            <div key={p.Id} draggable onDragStart={(e) => e.dataTransfer.setData("personId", p.Id.toString())} className={styles.draggablePersonaCard}>
              <Persona text={p.NombreyApellido} imageUrl={p.FotoPerfil} size={PersonaSize.size24} />
            </div>
          ))}
        </div>
      </div>

      <Stack horizontal tokens={{ childrenGap: 15 }} styles={{ root: { width: "100%", alignItems: "start" } }}>
        <div className={styles.tableContainer}>
          <table className={styles.planTable}>
            <thead>
              <tr>
                <th className={styles.colObra}>Obra</th>
                {DIAS_SEMANA.map((d, i) => (
                  <th key={d} className={styles.colDia}>
                    <Stack horizontalAlign="center">
                      <Text variant="medium">{d}</Text>
                      <Text variant="small" style={{ fontSize: '10px', color: '#666' }}>
                        {getFechaPorDiaIndex(i).getDate()}/{getFechaPorDiaIndex(i).getMonth() + 1}
                      </Text>
                    </Stack>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {obras.map(obra => (
                <tr key={obra.Id}>
                  <td className={styles.cellObra}>
                    <Stack tokens={{ childrenGap: 4 }}>
                      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>{obra.Title}</Text>
                      <Text variant="small" styles={{ root: { color: '#666', fontSize: '11px' } }}>
                        Avance: {obra.ProgresoReal || 0}% • {obra.EstadoObra}
                      </Text>
                      <ProgressIndicator percentComplete={(obra.ProgresoReal || 0) / 100} styles={{ itemProgress: { padding: 0 }, progressBar: { backgroundColor: '#107c41' } }} />
                    </Stack>
                  </td>
                  {DIAS_SEMANA.map((_, i) => {
                    const fechaDia = getFechaPorDiaIndex(i).toDateString();
                    const asigsEnDia = asignaciones.filter(a => a.ObraId === obra.Id && new Date(a.FechaInicio).toDateString() === fechaDia);
                    return (
                      <td key={i} onDragOver={e => e.preventDefault()} onDrop={e => onDrop(e, obra.Id, i)} className={styles.dropZone}>
                        <div className={styles.asignadosConsola}>
                          {asigsEnDia.map(a => {
                            const p = personalDisponible.find(pers => pers.Id === a.PersonalId);
                            return p ? (
                              <div 
                                key={a.Id} 
                                draggable 
                                onDragStart={(e) => e.dataTransfer.setData("assignmentId", a.Id!.toString())} // <--- NUEVO: Enviar ID asignación para moverla
                                onClick={() => setSelectedAsig({asig: a, persona: p})} 
                                className={styles.fotoAsignada}
                                style={{ cursor: 'grab' }}
                                title="Arrastra para mover o haz clic para eliminar"
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

        <div className={styles.pendingPanel}>
          <Text className={styles.panelTituloCompacto}>Pendientes</Text>
          <div className={styles.pendingList}>
            {obrasPendientes.map((op, idx) => (
              <div key={idx} className={styles.pendingItem}>
                <Stack horizontal horizontalAlign="space-between">
                  <Text className={styles.pendingName}>{op.nombre}</Text>
                  <IconButton iconProps={{ iconName: "Cancel" }} onClick={() => setObrasPendientes(obrasPendientes.filter((_, i) => i !== idx))} />
                </Stack>
                <Text className={styles.pendingReason}>{op.motivo}</Text>
              </div>
            ))}
          </div>
        </div>
      </Stack>

      {/* --- DIALOGO PARA ELIMINAR ASIGNACIÓN --- */}
      <Dialog 
        hidden={!selectedAsig} 
        onDismiss={() => setSelectedAsig(null)} 
        dialogContentProps={{ 
          type: DialogType.normal, 
          title: "Gestionar Asignación",
          subText: selectedAsig ? `¿Estás seguro de que quieres eliminar la asignación de ${selectedAsig.persona.NombreyApellido} para este día?` : ""
        }}
      >
        <DialogFooter>
          <PrimaryButton 
            onClick={eliminarAsignacion} 
            text="Eliminar" 
            style={{ backgroundColor: '#d13438', borderColor: '#d13438' }} // Botón en rojo para indicar borrado
          />
          <DefaultButton onClick={() => setSelectedAsig(null)} text="Cancelar" />
        </DialogFooter>
      </Dialog>

      <Dialog hidden={!showAddPending} onDismiss={() => setShowAddPending(false)} dialogContentProps={{ type: DialogType.normal, title: "Nueva Nota Pendiente" }}>
        <TextField label="Nombre" value={newPending.nombre} onChange={(_, v) => setNewPending({ ...newPending, nombre: v || "" })} />
        <TextField label="Motivo" multiline rows={3} value={newPending.motivo} onChange={(_, v) => setNewPending({ ...newPending, motivo: v || "" })} />
        <DialogFooter>
          <PrimaryButton onClick={() => { setObrasPendientes([...obrasPendientes, newPending]); setNewPending({ nombre: "", motivo: "" }); setShowAddPending(false); }} text="Añadir" />
          <DefaultButton onClick={() => setShowAddPending(false)} text="Cancelar" />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};