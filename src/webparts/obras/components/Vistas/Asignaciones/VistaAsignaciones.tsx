import * as React from "react";
import {
    Stack,
    Text,
    Persona,
    PersonaSize,
    Dropdown,
    PrimaryButton,
    IconButton,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    DatePicker,
    Separator,
    Icon,
} from "@fluentui/react";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { IObra } from "../../../models/IObra";
import { IPersonal } from "../../../models/IPersonal";
import { IAsignacion } from "../../../models/IAsignacion";
import styles from "./VistaAsignaciones.module.scss";

export const VistaAsignaciones: React.FC<{ context: any }> = (props) => {
    const [data, setData] = React.useState<{
        obras: IObra[];
        personal: IPersonal[];
        asignaciones: IAsignacion[];
    }>({ obras: [], personal: [], asignaciones: [] });

    const [loading, setLoading] = React.useState(true);
    const [error, setError] = React.useState<string | null>(null);
    const [seleccion, setSeleccion] = React.useState({
        obraId: 0,
        personalId: 0,
        fechaFin: new Date(),
    });

    const service = React.useMemo(() => new AsignacionesService(props.context), [props.context]);

    const cargarDatos = React.useCallback(async () => {
        setLoading(true);
        try {
            const [obrasData, personalData, asignacionesData] = await Promise.all([
                service.getObrasActivas(),
                service.getPersonalDisponible(),
                service.getAsignaciones(),
            ]);
            setData({ obras: obrasData, personal: personalData, asignaciones: asignacionesData });
            setError(null);
        } catch (err) {
            setError("Error al cargar los datos de asignación.");
        } finally {
            setLoading(false);
        }
    }, [service]);

    React.useEffect(() => {
        cargarDatos();
    }, [cargarDatos]);

    const handleAsignar = async () => {
        if (!seleccion.obraId || !seleccion.personalId || !seleccion.fechaFin) {
            alert("Por favor complete todos los campos requeridos");
            return;
        }

        try {
            await service.crearAsignacion(seleccion.obraId, seleccion.personalId, seleccion.fechaFin);
            setSeleccion((prev) => ({ ...prev, personalId: 0 }));
            await cargarDatos();
        } catch (err) {
            alert("No se pudo registrar la asignación.");
        }
    };

    const handleEliminar = async (id: number) => {
        if (!confirm("¿Está seguro de remover este personal de la obra?")) return;
        try {
            await service.eliminarAsignacion(id);
            await cargarDatos();
        } catch (err) {
            alert("Error al eliminar la asignación.");
        }
    };

    const formatearFecha = (dateString?: string | Date): string => {
        if (!dateString) return new Date().toLocaleDateString();
        const d = new Date(dateString);
        return isNaN(d.getTime()) ? new Date().toLocaleDateString() : d.toLocaleDateString();
    };

    if (loading) return <Spinner size={SpinnerSize.large} label="Sincronizando cuadrillas..." />;

    return (
        <div className={styles.container}>
            <Text variant="xLarge" className={styles.title}>Gestión de Personal y Asignaciones</Text>
            <Text className={styles.subtitle}>Distribuye los operarios disponibles en los frentes de obra activos</Text>

            {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

            <div className={styles.panelAsignacion}>
                <Stack tokens={{ childrenGap: 15 }}>
                    <Text variant="large" className={styles.panelTitle}>Nueva Programación de Obra</Text>
                    
                    <Dropdown
                        label="Seleccionar Frente de Obra"
                        placeholder="Elija un proyecto activo"
                        selectedKey={seleccion.obraId || undefined}
                        options={data.obras.map((o) => ({ key: o.Id!, text: o.Title }))}
                        onChange={(_, opt) => setSeleccion((prev) => ({ ...prev, obraId: opt?.key as number }))}
                    />

                    <Dropdown
                        label="Seleccionar Operario"
                        placeholder="Elija un operario para el sitio"
                        selectedKey={seleccion.personalId || undefined}
                        options={data.personal.map((p) => ({ key: p.Id!, text: `${p.NombreyApellido} — (${p.Rol})` }))}
                        onChange={(_, opt) => setSeleccion((prev) => ({ ...prev, personalId: opt?.key as number }))}
                    />

                    <DatePicker
                        label="Fecha de Trabajo / Ejecución"
                        placeholder="¿Qué día asiste a la obra?"
                        value={seleccion.fechaFin}
                        onSelectDate={(date) => setSeleccion((prev) => ({ ...prev, fechaFin: date || new Date() }))}
                    />

                    <PrimaryButton
                        text="Asignar y Programar"
                        onClick={handleAsignar}
                        disabled={!seleccion.obraId || !seleccion.personalId}
                    />
                </Stack>
            </div>

            <Separator className={styles.separator} />

            <div className={styles.gridObras}>
                {data.obras.map((o) => {
                    const asignados = data.asignaciones.filter((a) => a.ObraId === o.Id);

                    return (
                        <div key={o.Id} className={styles.obraCard}>
                            <Stack tokens={{ childrenGap: 10 }}>
                                <div className={styles.obraHeader}>
                                    <Text className={styles.obraTitle}>{o.Title}</Text>
                                    <Text className={styles.obraUbicacion}>{o.DireccionObra}</Text>
                                </div>

                                <Text variant="medium" style={{ fontWeight: 600, marginTop: 5 }}>Cuadrilla Programada:</Text>
                                
                                <Stack tokens={{ childrenGap: 10 }} className={styles.personalList}>
                                    {asignados.length > 0 ? (
                                        asignados.map((asig) => {
                                            const p = data.personal.find((per) => per.Id === asig.PersonalId);
                                            const semaforo = service.calcularSemaforoAsignacion(asig.FechaFinPrevista?.toString());

                                            return (
                                                <div key={asig.Id} className={styles.personalRow}>
                                                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ width: "100%" }}>
                                                        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" style={{ flexGrow: 1 }}>
                                                            <Persona
                                                                imageUrl={p?.FotoPerfil}
                                                                size={PersonaSize.size40}
                                                                presence={semaforo.presence}
                                                            />
                                                            <Stack tokens={{ childrenGap: 2 }}>
                                                                <Text className={styles.personaName}>{p?.NombreyApellido}</Text>
                                                                <Text className={styles.semaforoText} style={{ color: semaforo.presence === 4 ? "#d83b01" : "#107c41" }}>
                                                                    {semaforo.label}
                                                                </Text>
                                                                
                                                                <Stack horizontal tokens={{ childrenGap: 15 }} style={{ marginTop: 4, opacity: 0.85 }}>
                                                                    <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                                                                        <Icon iconName="CalendarSettings" style={{ fontSize: 12, color: "#0078d4" }} />
                                                                        <Text variant="small" style={{ color: "#323130" }}>
                                                                            <b>Prog:</b> {formatearFecha((asig as any).Created)}
                                                                        </Text>
                                                                    </Stack>
                                                                    <Stack horizontal tokens={{ childrenGap: 4 }} verticalAlign="center">
                                                                        <Icon iconName="Balloons" style={{ fontSize: 12, color: "#107c41" }} />
                                                                        <Text variant="small" style={{ color: "#323130" }}>
                                                                            <b>Trabajo:</b> {formatearFecha(asig.FechaFinPrevista)}
                                                                        </Text>
                                                                    </Stack>
                                                                </Stack>
                                                            </Stack>
                                                        </Stack>
                                                        <IconButton
                                                            iconProps={{ iconName: "Cancel" }}
                                                            className={styles.deleteBtn}
                                                            onClick={() => handleEliminar(asig.Id!)}
                                                        />
                                                    </Stack>
                                                </div>
                                            );
                                        })
                                    ) : (
                                        <Text className={styles.emptyText}>Sin personal asignado actualmente</Text>
                                    )}
                                </Stack>
                            </Stack>
                        </div>
                    );
                })}
            </div>
        </div>
    );
};