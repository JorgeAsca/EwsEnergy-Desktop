import * as React from "react";
import {
  Stack, Text, Spinner, SpinnerSize, MessageBar, 
  MessageBarType, PrimaryButton, DefaultButton, TextField, 
  Separator, IconButton, Dialog, DialogType, DialogFooter, Modal,
  DetailsList, DetailsListLayoutMode, SelectionMode, IColumn
} from "@fluentui/react";

// Importamos el servicio y el modelo
import { ClientesService } from "../../../service/ClientesService"; 
import { ICliente } from "../../../models/ICliente";
import styles from "./VistaClientes.module.scss";

// Cambiamos SPFI por el context nativo de SharePoint que usas en toda la app
export const ListaClientes: React.FC<{ context: any }> = (props) => {
  const [clientes, setClientes] = React.useState<ICliente[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [mensaje, setMensaje] = React.useState<{ texto: string; tipo: MessageBarType } | null>(null);
  const [isOpen, setIsOpen] = React.useState(false);
  const [procesando, setProcesando] = React.useState(false);
  const [editandoId, setEditandoId] = React.useState<number | null>(null);
  const [hideDeleteDialog, setHideDeleteDialog] = React.useState(true);

  const [formulario, setFormulario] = React.useState<Partial<ICliente>>({
    Title: "", CIF: "", Direccion: "", Email: "", Telefono: ""
  });

  // Instanciamos el servicio igual que en VistaAsignaciones o GaleriaPersonal
  const service = React.useMemo(() => new ClientesService(props.context), [props.context]);

  const cargarClientes = async () => {
    try {
      setLoading(true);
      const data = await service.getClientes();
      setClientes(data);
    } catch (error) {
      setMensaje({ texto: "Error al cargar los clientes", tipo: MessageBarType.error });
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => { 
    cargarClientes(); 
  }, [service]); // Dependencia actualizada al servicio

  const handleGuardar = async () => {
    if (!formulario.Title) {
      setMensaje({ texto: "El nombre es obligatorio", tipo: MessageBarType.warning });
      return;
    }
    try {
      setProcesando(true);
      if (editandoId) {
        await service.actualizarCliente(editandoId, formulario);
      } else {
        await service.crearCliente(formulario);
      }
      setIsOpen(false);
      await cargarClientes();
    } catch (error) {
      setMensaje({ texto: "Error al guardar.", tipo: MessageBarType.error });
    } finally { 
      setProcesando(false); 
    }
  };

  const handleEliminar = async () => {
    if (!editandoId) return;
    try {
      setProcesando(true);
      await service.eliminarCliente(editandoId);
      setHideDeleteDialog(true);
      setIsOpen(false);
      await cargarClientes();
    } catch (error) {
      setMensaje({ texto: "Error al eliminar.", tipo: MessageBarType.error });
    } finally { 
      setProcesando(false); 
    }
  };

  const abrirEditor = (cliente?: ICliente) => {
    if (cliente) {
      setEditandoId(cliente.Id || null);
      setFormulario(cliente);
    } else {
      setEditandoId(null);
      setFormulario({ Title: "", CIF: "", Direccion: "", Email: "", Telefono: "" });
    }
    setIsOpen(true);
  };

  const columns: IColumn[] = [
    {
      key: "colNombre", name: "Nombre", fieldName: "Title", minWidth: 140, maxWidth: 200,
      onRender: (item: ICliente) => <div data-label="Nombre:"><Text style={{ fontWeight: 600 }}>{item.Title}</Text></div>
    },
    { 
      key: "colCIF", name: "CIF", fieldName: "CIF", minWidth: 80, maxWidth: 90,
      onRender: (item: ICliente) => <div data-label="CIF:">{item.CIF}</div>
    },
    { 
      key: "colTelefono", name: "Teléfono", fieldName: "Telefono", minWidth: 90, maxWidth: 110,
      onRender: (item: ICliente) => <div data-label="Tel:">{item.Telefono}</div>
    },
    { 
      key: "colEmail", name: "Email", fieldName: "Email", minWidth: 130, maxWidth: 180,
      onRender: (item: ICliente) => (
        <div data-label="Email:">
            {item.Email ? <a href={`mailto:${item.Email}`} style={{ color: '#0078d4', textDecoration: 'none' }}>{item.Email}</a> : "-"}
        </div>
      )
    },
    {
      key: "colAcciones", name: "Acciones", minWidth: 80, maxWidth: 80,
      onRender: (item: ICliente) => (
        <div data-label="Opciones:">
            <Stack horizontal tokens={{ childrenGap: 5 }}>
                <IconButton iconProps={{ iconName: "Edit" }} onClick={() => abrirEditor(item)} />
                <IconButton iconProps={{ iconName: "Delete" }} onClick={() => { abrirEditor(item); setHideDeleteDialog(false); }} styles={{ root: { color: '#d13438' } }} />
            </Stack>
        </div>
      )
    }
  ];

  return (
    <div className={styles.container}>
      <Stack className={styles.headerSection} horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }}>
          <Text variant="xxLarge" style={{ fontWeight: 600 }}>Directorio de Clientes</Text>
          <div style={{ backgroundColor: '#e1dfdd', padding: '2px 10px', borderRadius: '12px', fontSize: '12px' }}>
            {clientes.length} registrados
          </div>
        </Stack>
        <PrimaryButton className={styles.btnNuevoCliente} iconProps={{ iconName: "AddFriend" }} text="Nuevo Cliente" onClick={() => abrirEditor()} />
      </Stack>

      {mensaje && <MessageBar messageBarType={mensaje.tipo} onDismiss={() => setMensaje(null)} style={{ marginBottom: 20 }}>{mensaje.texto}</MessageBar>}

      <div className={styles.tableContainer}>
        {loading ? <Spinner size={SpinnerSize.large} style={{ padding: '40px' }} /> : (
          <DetailsList 
            items={clientes} 
            columns={columns} 
            layoutMode={DetailsListLayoutMode.justified} 
            selectionMode={SelectionMode.none} 
          />
        )}
      </div>

      <Modal isOpen={isOpen} onDismiss={() => setIsOpen(false)} isBlocking={false}>
        <div className={styles.modalContent}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="xLarge" style={{ fontWeight: 600 }}>{editandoId ? "Editar Cliente" : "Nuevo Cliente"}</Text>
            <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => setIsOpen(false)} />
          </Stack>
          <Separator style={{ margin: '15px 0' }} />
          <Stack tokens={{ childrenGap: 15 }}>
            <TextField label="Nombre del Cliente" value={formulario.Title} onChange={(_, v) => setFormulario({...formulario, Title: v || ""})} required />
            <Stack horizontal tokens={{ childrenGap: 10 }}>
                <TextField label="CIF" style={{ flex: 1 }} value={formulario.CIF} onChange={(_, v) => setFormulario({...formulario, CIF: v || ""})} />
                <TextField label="Teléfono" style={{ flex: 1 }} value={formulario.Telefono} onChange={(_, v) => setFormulario({...formulario, Telefono: v || ""})} />
            </Stack>
            <TextField label="Correo Electrónico" value={formulario.Email} onChange={(_, v) => setFormulario({...formulario, Email: v || ""})} />
            <TextField label="Dirección" multiline rows={2} value={formulario.Direccion} onChange={(_, v) => setFormulario({...formulario, Direccion: v || ""})} />
          </Stack>
          <Stack horizontal tokens={{ childrenGap: 10 }} horizontalAlign="end" style={{ marginTop: '25px' }}>
            {procesando ? <Spinner size={SpinnerSize.medium} /> : (
              <>
                <PrimaryButton className={styles.btnNuevoCliente} text={editandoId ? "Actualizar" : "Registrar"} onClick={handleGuardar} />
                <DefaultButton text="Cancelar" onClick={() => setIsOpen(false)} />
              </>
            )}
          </Stack>
        </div>
      </Modal>

      <Dialog hidden={hideDeleteDialog} onDismiss={() => setHideDeleteDialog(true)} 
        dialogContentProps={{ type: DialogType.normal, title: 'Confirmar eliminación', subText: `¿Eliminar al cliente ${formulario.Title}?` }}>
        <DialogFooter>
          <PrimaryButton onClick={handleEliminar} text="Eliminar" styles={{ root: { backgroundColor: '#d13438', borderColor: '#d13438' } }} />
          <DefaultButton onClick={() => setHideDeleteDialog(true)} text="Cancelar" />
        </DialogFooter>
      </Dialog>
    </div>
  );
};