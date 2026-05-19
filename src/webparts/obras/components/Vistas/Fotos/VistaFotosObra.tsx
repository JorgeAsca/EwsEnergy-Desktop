import * as React from "react";
import {
    Stack,
    Text,
    Persona,
    PersonaSize,
    PrimaryButton,
    DefaultButton,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    TextField,
    Icon,
    IconButton,
    Dropdown,
    IDropdownOption,
    Slider,
} from "@fluentui/react";
import { IPersonal } from "../../../models/IPersonal";
import { IObra } from "../../../models/IObra";
import { PersonalService } from "../../../service/PersonalService";
import { AsignacionesService } from "../../../service/AsignacionesService";
import { ProjectService } from "../../../service/ProjectService";
import { PhotoService } from "../../../service/PhotoService";
import styles from "./VistaFotosObra.module.scss";

interface IHorasPersonal {
    [key: number]: number; // ID Personal -> Horas dedicadas
}

export const VistaFotosObra: React.FC<{ context: any }> = (props) => {
    const [paso, setPaso] = React.useState(1);
    const [loading, setLoading] = React.useState(true);
    const [subiendo, setSubiendo] = React.useState(false);

    // Inputs independientes para fotos
    const fileInputRefFinal = React.useRef<HTMLInputElement>(null);
    const fileInputRefPrevia = React.useRef<HTMLInputElement>(null);
    
    const [mensajeExito, setMensajeExito] = React.useState(false);
    const [procesandoCaptura, setProcesandoCaptura] = React.useState(false);

    const [operario, setOperario] = React.useState<IPersonal | null>(null);
    const [obraSeleccionada, setObraSeleccionada] = React.useState<IObra | null>(null);

    // Gestión de cuadrilla y horas
    const [compañeros, setCompañeros] = React.useState<IPersonal[]>([]);
    const [horasTrabajadas, setHorasTrabajadas] = React.useState<IHorasPersonal>({});

    // Listados de fotos separados
    const [fotosPrevias, setFotosPrevias] = React.useState<any[]>([]);
    const [fotosFinales, setFotosFinales] = React.useState<any[]>([]);
    const [comentarios, setComentarios] = React.useState("");

    const [data, setData] = React.useState({
        listaPersonal: [] as IPersonal[],
        obrasActivas: [] as IObra[],
        asignacionesGlobales: [] as any[]
    });

    const services = React.useMemo(() => ({
        personal: new PersonalService(props.context),
        asig: new AsignacionesService(props.context),
        obras: new ProjectService(props.context),
        fotos: new PhotoService(props.context),
    }), [props.context]);

    React.useEffect(() => {
        const iniciar = async () => {
            try {
                const [pers, asigs, obs] = await Promise.all([
                    services.personal.getPersonal(),
                    services.asig.getAsignaciones(),
                    services.obras.getObras(),
                ]);
                
                const obrasActivasFiltradas = obs.filter((o) => o.EstadoObra !== "Finalizado");

                setData({
                    listaPersonal: pers,
                    asignacionesGlobales: asigs,
                    obrasActivas: obrasActivasFiltradas
                });

                const currentUserEmail = props.context.pageContext.user.email.toLowerCase();
                const yoMismo = pers.find(p => p.Email && p.Email.toLowerCase() === currentUserEmail);
                
                if (yoMismo) {
                    setOperario(yoMismo);
                    // Inicializamos las horas del operario actual por defecto en 8h (100%)
                    if (yoMismo.Id) {
                        setHorasTrabajadas(prev => ({ ...prev, [yoMismo.Id as number]: 8 }));
                    }
                }

            } catch (e) {
                console.error(e);
            } finally {
                setLoading(false);
            }
        };
        iniciar();
    }, [services]);

    const handleSeleccionarObra = (ob: IObra) => {
        setObraSeleccionada(ob);
        
        // Cargar los compañeros asignados automáticamente a esta obra
        const asigsObra = data.asignacionesGlobales.filter(a => Number(a.ObraId) === Number(ob.Id));
        const compis = data.listaPersonal.filter(p =>
            asigsObra.some(a => Number(a.PersonalId) === Number(p.Id)) && p.Id !== operario?.Id
        );
        
        setCompañeros(compis);

        // Inicializar las horas de los compañeros asignados por defecto en 8 horas
        const horasIniciales: IHorasPersonal = {};
        if (operario?.Id) horasIniciales[operario.Id as number] = horasTrabajadas[operario.Id as number] || 8;
        compis.forEach(c => {
            if (c.Id) horasIniciales[c.Id as number] = 8;
        });
        setHorasTrabajadas(horasIniciales);

        setPaso(2);
    };

    const agregarCompañeroExtra = (event: any, option?: IDropdownOption) => {
        if (option) {
            const persona = data.listaPersonal.find(p => p.Id === option.key);
            if (persona && persona.Id) {
                setCompañeros(prev => [...prev, persona]);
                setHorasTrabajadas(prev => ({ ...prev, [persona.Id as number]: 8 }));
            }
        }
    };

    const removerCompañero = (id: number) => {
        setCompañeros(prev => prev.filter(c => c.Id !== id));
        setHorasTrabajadas(prev => {
            const copia = { ...prev };
            delete copia[id];
            return copia;
        });
    };

    const cambiarHoras = (id: number, nuevasHoras: number) => {
        setHorasTrabajadas(prev => ({ ...prev, [id]: nuevasHoras }));
    };

    const manejarCapturaFoto = async (event: React.ChangeEvent<HTMLInputElement>, esPrevia: boolean) => {
        const archivo = event.target.files?.[0];
        if (!archivo) return;

        setProcesandoCaptura(true);
        setMensajeExito(false);

        try {
            const ubicacion = await obtenerUbicacion();
            const nuevaFotoLocal = {
                ID: Date.now(),
                archivo: archivo,
                Url: URL.createObjectURL(archivo),
                Nombre: archivo.name,
                latitud: ubicacion?.lat,
                longitud: ubicacion?.lng,
                Ubicacion: ubicacion ? `${ubicacion.lat}, ${ubicacion.lng}` : "Capturada"
            };

            if (esPrevia) {
                setFotosPrevias((prev) => [...prev, nuevaFotoLocal]);
            } else {
                setFotosFinales((prev) => [...prev, nuevaFotoLocal]);
            }
            
            setMensajeExito(true);
            setTimeout(() => setMensajeExito(false), 3000);
        } catch (error) {
            console.error("Error en vista previa:", error);
        } finally {
            setProcesandoCaptura(false);
            if (fileInputRefPrevia.current) fileInputRefPrevia.current.value = "";
            if (fileInputRefFinal.current) fileInputRefFinal.current.value = "";
        }
    };

    const obtenerUbicacion = (): Promise<{ lat: number; lng: number } | null> => {
        return new Promise((resolve) => {
            if (!navigator.geolocation) resolve(null);
            navigator.geolocation.getCurrentPosition(
                (pos) => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
                () => resolve(null),
                { enableHighAccuracy: true, timeout: 5000 }
            );
        });
    };

    const enviarReporte = async () => {
        if (!obraSeleccionada || !operario || fotosFinales.length === 0) return;

        setSubiendo(true);
        try {
            // 1. Subir fotos previas
            for (const fotoObj of fotosPrevias) {
                await services.fotos.uploadCompressedPhoto(fotoObj.archivo, `${obraSeleccionada.Title}_Previas`, {
                    operario: operario.NombreyApellido,
                    operarioId: operario.Id as number,
                    obraId: obraSeleccionada.Id as number,
                    comentarios: "Registro previo de entrada",
                    latitud: fotoObj.latitud,
                    longitud: fotoObj.longitud
                });
            }

            // 2. Subir fotos finales de cierre
            for (const fotoObj of fotosFinales) {
                await services.fotos.uploadCompressedPhoto(fotoObj.archivo, obraSeleccionada.Title, {
                    operario: operario.NombreyApellido,
                    operarioId: operario.Id as number,
                    obraId: obraSeleccionada.Id as number,
                    comentarios: comentarios,
                    latitud: fotoObj.latitud,
                    longitud: fotoObj.longitud
                });
            }

            // 3. Calcular consumo de jornadas (8h trabajadas = 1 jornada)
            let totalHorasCuadrilla = 0;
            Object.keys(horasTrabajadas).forEach(key => {
                totalHorasCuadrilla += horasTrabajadas[Number(key)] || 0;
            });
            
            const jornadasConsumidasHoy = totalHorasCuadrilla / 8;
            
            // 4. Descontar las jornadas consumidas del total del proyecto
            const jornadasTotales = obraSeleccionada.JornadasTotales || 30;
            const nuevoProgresoReal = Math.min((obraSeleccionada.ProgresoReal || 0) + ((jornadasConsumidasHoy / jornadasTotales) * 100), 100);

            // Actualizamos en SharePoint el progreso calculado de forma automática
            await services.obras.actualizarProgresoObra(obraSeleccionada.Id as number, parseFloat(nuevoProgresoReal.toFixed(2)));

            alert(`¡Reporte enviado! Se registraron ${totalHorasCuadrilla}h en total (${jornadasConsumidasHoy.toFixed(2)} jornadas descontadas del proyecto).`);
            window.location.reload();
        } catch (error) {
            alert("Hubo un error al sincronizar las evidencias de la obra.");
        } finally {
            setSubiendo(false);
        }
    };

    if (loading) return <Spinner size={SpinnerSize.large} label="Cargando información del proyecto..." />;

    const direccionCodificada = obraSeleccionada ? encodeURIComponent(obraSeleccionada.DireccionObra || obraSeleccionada.Title) : "";
    const urlMapaInteractivo = `https://maps.google.com/maps?q=${direccionCodificada}&t=&z=15&ie=UTF8&iwloc=&output=embed`;

    return (
        <div className={styles.container}>
            <header className={styles.appHeader}>
                <Stack>
                    <Text variant="xLarge" className={styles.title}>EWS</Text>
                    <Text className={styles.subtitle}>Portal de Obra Inteligente</Text>
                </Stack>
                {operario && <Persona imageUrl={operario.FotoPerfil} size={PersonaSize.size32} hidePersonaDetails />}
            </header>

            {/* Indicador visual de los 5 pasos actuales */}
            <div className={styles.wizardNav}>
                {[1, 2, 3, 4, 5].map((p) => (
                    <div key={p} className={`${styles.dot} ${paso >= p ? styles.active : ""}`} />
                ))}
            </div>

            <main className={styles.mainContent}>
                {/* PASO 1: SELECCIÓN DE OBRA */}
                {paso === 1 && (
                    <section className={styles.stepContainer}>
                        <Text variant="large" className={styles.stepTitle}>1. Selección de Obra</Text>
                        {!operario ? (
                            <Stack tokens={{ childrenGap: 10 }}>
                                <Text>¿Quién envía el reporte?</Text>
                                {data.listaPersonal.map((p) => (
                                    <div key={p.Id} className={styles.userCard} onClick={() => setOperario(p)}>
                                        <Persona imageUrl={p.FotoPerfil} text={p.NombreyApellido} secondaryText={p.Rol} size={PersonaSize.size40} />
                                    </div>
                                ))}
                            </Stack>
                        ) : (
                            <Stack tokens={{ childrenGap: 15 }}>
                                <Text>Hola <b>{operario.NombreyApellido}</b>, selecciona una obra activa:</Text>
                                {data.obrasActivas.map((o) => (
                                    <div key={o.Id} className={styles.obraCard} onClick={() => handleSeleccionarObra(o)}>
                                        <Text className={styles.obraTitle}>{o.Title}</Text>
                                        <Text variant="small"><Icon iconName="MapPin" /> {o.DireccionObra}</Text>
                                    </div>
                                ))}
                                <DefaultButton text="Cambiar Operario" onClick={() => setOperario(null)} style={{ marginTop: 10 }} />
                            </Stack>
                        )}
                    </section>
                )}

                {/* PASO 2: DATOS DE LA OBRA */}
                {paso === 2 && obraSeleccionada && (
                    <section className={styles.stepContainer}>
                        <Text variant="large" className={styles.stepTitle}>2. Datos de la Obra</Text>
                        
                        <div style={{ background: '#f3f2f1', padding: '15px', borderRadius: '8px', marginBottom: '20px' }}>
                            <Text variant="mediumPlus" style={{ fontWeight: 'bold', display: 'block', marginBottom: '5px' }}>
                                {obraSeleccionada.Title}
                            </Text>
                            <Text variant="small" style={{ color: '#605e5c', display: 'block', marginBottom: '15px' }}>
                                <Icon iconName="MapPin" style={{ marginRight: '5px' }} /> {obraSeleccionada.DireccionObra || "Dirección no especificada"}
                            </Text>

                            <div style={{ width: '100%', height: '220px', borderRadius: '6px', overflow: 'hidden', boxShadow: '0 2px 4px rgba(0,0,0,0.1)', marginBottom: '15px', border: '1px solid #ced4da' }}>
                                <iframe width="100%" height="100%" style={{ border: 0 }} src={urlMapaInteractivo} loading="lazy" />
                            </div>

                            <Stack tokens={{ childrenGap: 12 }}>
                                <div style={{ background: '#ffffff', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #107c41', display: 'flex', alignItems: 'center' }}>
                                    <Icon iconName="BuildDefinition" style={{ fontSize: '20px', color: '#107c41', marginRight: '10px' }} />
                                    <Stack style={{ flexGrow: 1 }}>
                                        <Text style={{ fontWeight: '600' }}>Planos Técnicos</Text>
                                        <Text variant="small" style={{ color: '#a19f9d' }}>Esquemas estructurales y eléctricos</Text>
                                    </Stack>
                                    <IconButton iconProps={{ iconName: "DietPlanView" }} title="Ver Planos" onClick={() => alert("Módulo de planos (Próximamente)...")} />
                                </div>

                                <div style={{ background: '#ffffff', padding: '12px', borderRadius: '6px', borderLeft: '4px solid #d83b01', display: 'flex', alignItems: 'center' }}>
                                    <Icon iconName="PDF" style={{ fontSize: '20px', color: '#d83b01', marginRight: '10px' }} />
                                    <Stack style={{ flexGrow: 1 }}>
                                        <Text style={{ fontWeight: '600' }}>Documentación y Permisos</Text>
                                        <Text variant="small" style={{ color: '#a19f9d' }}>Hojas de seguridad y actas</Text>
                                    </Stack>
                                    <IconButton iconProps={{ iconName: "DocumentSearch" }} title="Ver Documentos" onClick={() => alert("Biblioteca de documentos (Próximamente)...")} />
                                </div>
                            </Stack>
                        </div>

                        <Stack horizontal tokens={{ childrenGap: 10 }}>
                            <DefaultButton text="Atrás" onClick={() => setPaso(1)} />
                            <PrimaryButton text="Tomar fotos previas" onClick={() => setPaso(3)} className={styles.btnEws} style={{ flex: 1 }} />
                        </Stack>
                    </section>
                )}

                {/* PASO 3: FOTOS PREVIAS (LLEGADA) */}
                {paso === 3 && obraSeleccionada && (
                    <section className={styles.stepContainer}>
                        <Text variant="large" className={styles.stepTitle}>3. Fotos Previas (Llegada)</Text>
                        <input type="file" accept="image/*" capture="environment" style={{ display: "none" }} ref={fileInputRefPrevia} onChange={(e) => manejarCapturaFoto(e, true)} />

                        <label
                            className={styles.photoDropzone}
                            onClick={() => !procesandoCaptura && fileInputRefPrevia.current?.click()}
                            style={{ cursor: procesandoCaptura ? 'wait' : 'pointer', opacity: procesandoCaptura ? 0.7 : 1, border: '2px dashed #0078d4' }}
                        >
                            {procesandoCaptura ? (
                                <Spinner size={SpinnerSize.large} label="Registrando llegada..." />
                            ) : (
                                <><Icon iconName="Camera" className={styles.bigIcon} style={{ color: '#0078d4' }} /><Text>Capturar Estado Inicial de la Obra (GPS)</Text></>
                            )}
                        </label>

                        <div className={styles.previewContainer}>
                            {fotosPrevias.map((f, i) => (
                                <div key={f.ID || i} className={styles.previewItem}>
                                    <img src={f.Url} alt="preview previa" />
                                    {f.latitud && <span className={styles.gpsBadge}><Icon iconName="MapPin" /></span>}
                                    <IconButton iconProps={{ iconName: "Cancel" }} className={styles.removePhoto} onClick={() => setFotosPrevias(prev => prev.filter((_, idx) => idx !== i))} />
                                </div>
                            ))}
                        </div>

                        <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 25 } }}>
                            <DefaultButton text="Atrás" onClick={() => setPaso(2)} disabled={procesandoCaptura} />
                            <PrimaryButton 
                                text="Gestionar Personal" 
                                onClick={() => setPaso(4)} 
                                disabled={fotosPrevias.length === 0 || procesandoCaptura} 
                                className={styles.btnEws} 
                                style={{ flex: 1 }} 
                            />
                        </Stack>
                    </section>
                )}

                {/* PASO 4: GESTIONAR PERSONAL */}
                {paso === 4 && obraSeleccionada && (
                    <section className={styles.stepContainer}>
                        <Text variant="large" className={styles.stepTitle}>4. Gestionar Personal</Text>
                        <Text style={{ display: 'block', marginBottom: '15px', color: '#605e5c' }}>
                            Ajusta las horas de trabajo del personal. <b>Máximo 8 Horas por jornada (100%)</b>.
                        </Text>

                        <div style={{ background: '#f3f2f1', padding: '15px', borderRadius: '8px', marginBottom: '15px' }}>
                            <Stack tokens={{ childrenGap: 20 }}>
                                
                                {/* 1. Control del Operario Principal (Tope en max={8}) */}
                                {operario && operario.Id && (
                                    <div style={{ background: '#ffffff', padding: '12px', borderRadius: '6px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                                        <Persona imageUrl={operario.FotoPerfil} text={`${operario.NombreyApellido} (Tú)`} secondaryText={operario.Rol} size={PersonaSize.size32} />
                                        <div style={{ marginTop: '10px' }}>
                                            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '2px' }}>
                                                <Text variant="small" style={{ fontWeight: '600', color: '#0078d4' }}>
                                                    Horas: {horasTrabajadas[operario.Id as number] || 0}h
                                                </Text>
                                                <Text variant="small" style={{ fontWeight: '600', color: '#0078d4' }}>
                                                    {Math.round(((horasTrabajadas[operario.Id as number] || 0) / 8) * 100)}%
                                                </Text>
                                            </div>
                                            <Slider min={0} max={8} step={0.5} value={horasTrabajadas[operario.Id as number] || 0} showValue={false} onChange={(v) => cambiarHoras(operario.Id as number, v)} />
                                        </div>
                                    </div>
                                )}

                                {/* 2. Control de los Compañeros de Cuadrilla (Tope en max={8}) */}
                                {compañeros.map((c) => {
                                    if (!c.Id) return null;
                                    const hrs = horasTrabajadas[c.Id as number] || 0;
                                    const pct = Math.round((hrs / 8) * 100);
                                    return (
                                        <div key={c.Id} style={{ background: '#ffffff', padding: '12px', borderRadius: '6px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', position: 'relative' }}>
                                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                                <Persona imageUrl={c.FotoPerfil} text={c.NombreyApellido} secondaryText={c.Rol} size={PersonaSize.size32} />
                                                <IconButton iconProps={{ iconName: "Delete" }} title="Quitar de la lista" styles={{ root: { color: '#a19f9d' }, rootHovered: { color: '#d83b01' } }} onClick={() => removerCompañero(c.Id as number)} />
                                            </div>
                                            <div style={{ marginTop: '10px' }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '2px' }}>
                                                    <Text variant="small" style={{ fontWeight: '600', color: '#107c41' }}>Horas: {hrs}h</Text>
                                                    <Text variant="small" style={{ fontWeight: '600', color: '#107c41' }}>{pct}%</Text>
                                                </div>
                                                <Slider min={0} max={8} step={0.5} value={hrs} showValue={false} onChange={(v) => cambiarHoras(c.Id as number, v)} />
                                            </div>
                                        </div>
                                    );
                                })}
                            </Stack>
                        </div>

                        {/* Dropdown para añadir personal por imprevistos */}
                        <div style={{ marginBottom: '20px' }}>
                            <Dropdown
                                placeholder="+ Añadir personal por imprevisto"
                                options={data.listaPersonal.filter(p => p.Id !== operario?.Id && !compañeros.some(c => c.Id === p.Id)).map(p => ({ key: p.Id as number, text: p.NombreyApellido }))}
                                onChange={agregarCompañeroExtra}
                                selectedKey={null}
                            />
                        </div>

                        <Stack horizontal tokens={{ childrenGap: 10 }}>
                            <DefaultButton text="Atrás" onClick={() => setPaso(3)} />
                            <PrimaryButton text="Siguiente (Evidencias de Cierre)" onClick={() => setPaso(5)} className={styles.btnEws} style={{ flex: 1 }} />
                        </Stack>
                    </section>
                )}

                {/* PASO 5: EVIDENCIA VISUAL FINAL */}
                {paso === 5 && obraSeleccionada && (
                    <section className={styles.stepContainer}>
                        <Text variant="large" className={styles.stepTitle}>5. Evidencia Visual (Cierre)</Text>
                        <input type="file" accept="image/*" capture="environment" style={{ display: "none" }} ref={fileInputRefFinal} onChange={(e) => manejarCapturaFoto(e, false)} />

                        <label
                            className={styles.photoDropzone}
                            onClick={() => !procesandoCaptura && fileInputRefFinal.current?.click()}
                            style={{ cursor: procesandoCaptura ? 'wait' : 'pointer', opacity: procesandoCaptura ? 0.7 : 1 }}
                        >
                            {procesandoCaptura ? (
                                <Spinner size={SpinnerSize.large} label="Optimizando..." />
                            ) : (
                                <><Icon iconName="Camera" className={styles.bigIcon} /><Text>Toca para tomar foto de fin de jornada</Text></>
                            )}
                        </label>

                        <div className={styles.previewContainer}>
                            {fotosFinales.map((f, i) => (
                                <div key={f.ID || i} className={styles.previewItem}>
                                    <img src={f.Url} alt="preview final" />
                                    {f.latitud && <span className={styles.gpsBadge}><Icon iconName="MapPin" /></span>}
                                    <IconButton iconProps={{ iconName: "Cancel" }} className={styles.removePhoto} onClick={() => setFotosFinales(prev => prev.filter((_, idx) => idx !== i))} />
                                </div>
                            ))}
                        </div>

                        <TextField label="Comentarios de Cierre 🎤" multiline rows={3} value={comentarios} onChange={(_, v) => setComentarios(v || "")} />

                        <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 25 } }}>
                            <DefaultButton text="Atrás" onClick={() => setPaso(4)} disabled={subiendo || procesandoCaptura} />
                            <PrimaryButton 
                                text={subiendo ? "Sincronizando..." : "Enviar Reporte"} 
                                iconProps={{ iconName: "Send" }} 
                                onClick={enviarReporte} 
                                disabled={fotosFinales.length === 0 || subiendo || procesandoCaptura} 
                                className={styles.btnEws} 
                                style={{ flex: 1 }} 
                            />
                        </Stack>
                    </section>
                )}
            </main>
        </div>
    );
};