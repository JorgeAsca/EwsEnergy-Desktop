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

export const VistaPlanificacion: React.FC<{ context: WebPartContext }> = ({
  context,
}) => {
  const [obras, setObras] = React.useState<IObra[]>([]);
  const [personalDisponible, setPersonalDisponible] = React.useState<
    IPersonal[]
  >([]);
  const [asignaciones, setAsignaciones] = React.useState<IAsignacion[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [selectedAsig, setSelectedAsig] = React.useState<{
    asig: IAsignacion;
    persona: IPersonal;
  } | null>(null);
  const [obrasPendientes, setObrasPendientes] = React.useState<
    IObraPendiente[]
  >([]);
  const [showAddPending, setShowAddPending] = React.useState(false);
  const [newPending, setNewPending] = React.useState<IObraPendiente>({
    nombre: "",
    motivo: "",
  });

  const services = React.useMemo(
    () => ({
      project: new ProjectService(context),
      personal: new PersonalService(context),
      asig: new AsignacionesService(context),
    }),
    [context],
  );

  const getFechaPorDia = (nombreDia: string): Date => {
    const hoy = new Date();
    const lunes = new Date(
      hoy.setDate(hoy.getDate() - (hoy.getDay() || 7) + 1),
    );
    const index = DIAS_SEMANA.indexOf(nombreDia);
    const fechaResultado = new Date(lunes);
    fechaResultado.setDate(lunes.getDate() + index);
    return fechaResultado;
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

  React.useEffect(() => {
    cargarDatos();
  }, []);

  const onDrop = async (ev: React.DragEvent, obraId: number, dia: string) => {
    ev.preventDefault();
    const personId = parseInt(ev.dataTransfer.getData("personId"));
    const fecha = getFechaPorDia(dia);
    await services.asig.asignarPersonal({
      ObraId: obraId,
      PersonalId: personId,
      FechaInicio: fecha,
      FechaFinPrevista: fecha,
      EstadoProgreso: 0,
    });
    await cargarDatos();
  };

  const eliminarAsignacion = async () => {
    if (!selectedAsig?.asig.Id) return;
    await services.asig.eliminarAsignacion(selectedAsig.asig.Id);
    setSelectedAsig(null);
    await cargarDatos();
  };

  if (loading)
    return (
      <Spinner label="Cargando planificación..." size={SpinnerSize.large} />
    );

  return (
    <Stack tokens={{ childrenGap: 15 }} className={styles.vistaPlanificacion}>
      {/* HEADER */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" className={styles.titulo}>
          Planificación Semanal
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
              onDragStart={(e) =>
                e.dataTransfer.setData("personId", p.Id.toString())
              }
              className={styles.draggablePersonaCard}
            >
              <Persona
                text={p.NombreyApellido}
                imageUrl={p.FotoPerfil}
                size={PersonaSize.size24}
              />
            </div>
          ))}
        </div>
      </div>

      {/* CUERPO: TABLA Y PENDIENTES */}
      <Stack
        horizontal
        tokens={{ childrenGap: 15 }}
        styles={{ root: { width: "100%", alignItems: "start" } }}
      >
        <div className={styles.tableContainer}>
          <table className={styles.planTable}>
            <thead>
              <tr>
                <th className={styles.colObra}>Obra</th>
                {DIAS_SEMANA.map((d) => (
                  <th key={d} className={styles.colDia}>
                    {d}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {obras.map(obra => (
                <tr key={obra.Id}>
                  
                  {/* 1. COLUMNA DE LA OBRA (Aquí añadimos el Estado y el Progreso) */}
                  <td className={styles.cellObra}>
                    <Stack tokens={{ childrenGap: 4 }}>
                      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                        {obra.Title}
                      </Text>
                      <Stack>
                        {/* Mostramos el % de avance y el Estado (Ej: En Proceso) */}
                        <Text variant="small" styles={{ root: { color: '#666', fontSize: '11px' } }}>
                          Avance: {obra.ProgresoReal || 0}% • {obra.EstadoObra}
                        </Text>
                        <ProgressIndicator 
                          percentComplete={(obra.ProgresoReal || 0) / 100} 
                          styles={{ 
                            itemProgress: { padding: 0 }, 
                            progressBar: { backgroundColor: '#107c41' } 
                          }} 
                        />
                      </Stack>
                    </Stack>
                  </td>

                  {/* 2. COLUMNAS DE LOS DÍAS (Mantenemos intacto el Drag & Drop) */}
                  {DIAS_SEMANA.map(dia => {
                    const fechaDia = getFechaPorDia(dia).toDateString();
                    const asigsEnDia = asignaciones.filter(a => a.ObraId === obra.Id && new Date(a.FechaInicio).toDateString() === fechaDia);
                    return (
                      <td key={dia} onDragOver={e => e.preventDefault()} onDrop={e => onDrop(e, obra.Id, dia)} className={styles.dropZone}>
                        <div className={styles.asignadosConsola}>
                          {asigsEnDia.map(a => {
                            const p = personalDisponible.find(pers => pers.Id === a.PersonalId);
                            return p ? (
                              <div key={a.Id} onClick={() => setSelectedAsig({asig: a, persona: p})} className={styles.fotoAsignada}>
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

        {/* PANEL PENDIENTES CON SCROLL INDEPENDIENTE */}
        <div className={styles.pendingPanel}>
          <Text className={styles.panelTituloCompacto}>Pendientes</Text>
          <div className={styles.pendingList}>
            {obrasPendientes.length === 0 && (
              <span className={styles.emptyText}>Sin notas</span>
            )}
            {obrasPendientes.map((op, idx) => (
              <div key={idx} className={styles.pendingItem}>
                <Stack horizontal horizontalAlign="space-between">
                  <Text className={styles.pendingName}>{op.nombre}</Text>
                  <IconButton
                    iconProps={{ iconName: "Cancel" }}
                    styles={{ root: { height: 16, width: 16, fontSize: 10 } }}
                    onClick={() =>
                      setObrasPendientes(
                        obrasPendientes.filter((_, i) => i !== idx),
                      )
                    }
                  />
                </Stack>
                <Text className={styles.pendingReason}>{op.motivo}</Text>
              </div>
            ))}
          </div>
        </div>
      </Stack>

      {/* DIALOGS */}
      <Dialog
        hidden={!showAddPending}
        onDismiss={() => setShowAddPending(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Nueva Nota Pendiente",
        }}
      >
        <TextField
          label="Nombre"
          value={newPending.nombre}
          onChange={(_, v) => setNewPending({ ...newPending, nombre: v || "" })}
        />
        <TextField
          label="Motivo"
          multiline
          rows={3}
          value={newPending.motivo}
          onChange={(_, v) => setNewPending({ ...newPending, motivo: v || "" })}
        />
        <DialogFooter>
          <PrimaryButton
            onClick={() => {
              setObrasPendientes([...obrasPendientes, newPending]);
              setNewPending({ nombre: "", motivo: "" });
              setShowAddPending(false);
            }}
            text="Añadir"
          />
          <DefaultButton
            onClick={() => setShowAddPending(false)}
            text="Cancelar"
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!selectedAsig}
        onDismiss={() => setSelectedAsig(null)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Gestionar Asignación",
        }}
      >
        <DialogFooter>
          <PrimaryButton onClick={eliminarAsignacion} text="Eliminar" />
          <DefaultButton
            onClick={() => setSelectedAsig(null)}
            text="Cancelar"
          />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};
