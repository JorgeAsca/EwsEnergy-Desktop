import * as React from 'react';
import { 
    Stack, 
    Text, 
    SearchBox, 
    Spinner, 
    Icon, 
    Image, 
    ImageFit, 
    MessageBar, 
    MessageBarType 
} from '@fluentui/react';
import { DailyReportService } from '../../../service/DailyReportService';
import { IReporteHistorial } from '../../../models/IReporteHistorial';
import styles from './VistaHistorialTarjetas.module.scss';

export const VistaHistorialTarjetas: React.FC<{ context: any }> = (props) => {
    // --- ESTADOS ---
    const [reportes, setReportes] = React.useState<IReporteHistorial[]>([]);
    const [filtrados, setFiltrados] = React.useState<IReporteHistorial[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [error, setError] = React.useState<string | null>(null);

    // --- SERVICIO ---
    const service = React.useMemo(() => new DailyReportService(props.context), [props.context]);

    // --- CARGA DE DATOS ---
    const cargarDatos = async () => {
        try {
            setLoading(true);
            setError(null);
            const data = await service.getHistorialGlobal();
            setReportes(data);
            setFiltrados(data);
        } catch (e) {
            setError("Error al cargar el historial de evidencias. Por favor, intente de nuevo.");
            console.error(e);
        } finally {
            setLoading(false);
        }
    };

    React.useEffect(() => { 
        cargarDatos().catch(console.error); 
    }, []);

    // --- LÓGICA DE FILTRADO ---
    const onFilter = (text: string) => {
        if (!text) {
            setFiltrados(reportes);
            return;
        }
        
        const busqueda = text.toLowerCase();
        const filtrado = reportes.filter(r => 
            
            (r.Title && r.Title.toLowerCase().indexOf(busqueda) > -1) || 
            (r.Comentarios && r.Comentarios.toLowerCase().indexOf(busqueda) > -1)
        );
        setFiltrados(filtrado);
    };

    if (loading) return <Spinner label="Consultando archivos EWS..." className={styles.loader} />;

    return (
        <div className={styles.container}>
            <Stack tokens={{ childrenGap: 25 }}>
                <div className={styles.headerSection}>
                    <Stack>
                        <Text variant="xxLarge" className={styles.titulo}>Historial de Evidencias</Text>
                        <Text variant="small" className={styles.subtitulo}>Registro fotográfico de operaciones en campo</Text>
                    </Stack>
                    <SearchBox 
                        placeholder="Buscar por obra o comentario..." 
                        onSearch={onFilter} 
                        onChange={(_, val) => onFilter(val || "")}
                        className={styles.searchBar}
                    />
                </div>

                {error && (
                    <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
                        {error}
                    </MessageBar>
                )}

                <div className={styles.cardGrid}>
                    {filtrados.length > 0 ? (
                        filtrados.map((item) => (
                            <div key={item.Id} className={styles.reporteCard}>
                                <div className={styles.cardHeader}>
                                    <Text className={styles.obraName}>{item.Title}</Text>
                                    <Text className={styles.fechaText}>
                                        {item.FechaRegistro ? new Date(item.FechaRegistro).toLocaleDateString() : 'S/F'}
                                    </Text>
                                </div>
                                
                                <div className={styles.imageContainer}>
                                    <Image 
                                        src={item.UrlFoto?.Url} 
                                        alt="Foto reporte" 
                                        height={200} 
                                        imageFit={ImageFit.cover} 
                                        className={styles.reporteImagen}
                                    />
                                </div>

                                <div className={styles.cardContent}>
                                    <div className={styles.comentarioBox}>
                                        <Text className={styles.comentarios}>
                                            {item.Comentarios ? `"${item.Comentarios}"` : "Sin observaciones técnicas"}
                                        </Text>
                                    </div>
                                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} className={styles.footerOperario}>
                                        <Icon iconName="Contact" className={styles.iconOperario} />
                                        <Text variant="small">ID Operario: <b>{item.OperarioId}</b></Text>
                                    </Stack>
                                </div>
                            </div>
                        ))
                    ) : (
                        !error && <Text variant="large" styles={{ root: { textAlign: 'center', marginTop: 20 } }}>No se encontraron evidencias.</Text>
                    )}
                </div>
            </Stack>
        </div>
    );
};