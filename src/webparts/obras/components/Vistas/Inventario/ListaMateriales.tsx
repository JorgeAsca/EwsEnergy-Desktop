import * as React from "react";
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  IconButton,
  Dropdown,
  IDropdownOption,
  SearchBox,
  Separator,
  Spinner,
  SpinnerSize,
  Icon,
  Panel,
  PanelType,
  DefaultButton,
} from "@fluentui/react";
import styles from "./ListaMateriales.module.scss";
import { StockService } from "../../../service/StockService";

const categorias: IDropdownOption[] = [
  { key: "Consumible", text: "Consumible" },
  { key: "Herramienta", text: "Herramienta" },
  { key: "Maquinaria", text: "Maquinaria" },
  { key: "EPIS", text: "EPIS" },
];

export const ListaMateriales: React.FC<{ context: any }> = (props) => {
  const [items, setItems] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [filterText, setFilterText] = React.useState("");
  const [filterCategory, setFilterCategory] = React.useState<string>("Todas");
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [selectedItem, setSelectedItem] = React.useState<any>(null);
  
  const [nuevo, setNuevo] = React.useState({
    nombre: "",
    stock: 0,
    stockMin: 0,
    cat: "Consumible",
    ImagenPreview: "", // <-- Para previsualizar la imagen antes de subirla
    file: null as File | null
  });

  const service = React.useMemo(() => new StockService(props.context), [props.context]);

  const cargarMateriales = async () => {
    setLoading(true);
    try {
      const data = await service.getInventario();
      setItems(data);
    } catch (e) {
      console.error("Error al cargar inventario", e);
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    cargarMateriales();
  }, []);

  // --- MANEJADOR DE ARCHIVOS (IMÁGENES) ---
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>, mode: 'nuevo' | 'edit') => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const base64 = e.target?.result as string;
        if (mode === 'nuevo') {
          setNuevo({ ...nuevo, ImagenPreview: base64, file: file });
        } else {
          setSelectedItem({ ...selectedItem, ImagenPreview: base64, file: file });
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const handleAdd = async () => {
    if (!nuevo.nombre) return;
    try {
      setLoading(true);
      // Le pasamos el objeto y el archivo (nuevo.file)
      await service.crearMaterial({
        Title: nuevo.nombre,
        StockActual: nuevo.stock,
        StockMinimo: nuevo.stockMin,
        Categoria: nuevo.cat
      }, nuevo.file);
      
      setNuevo({ nombre: "", stock: 0, stockMin: 0, cat: "Consumible", ImagenPreview: "", file: null });
      await cargarMateriales();
    } catch (e) {
      console.error(e);
      alert("Error al crear el material. Revisa que los campos coincidan con SharePoint.");
    } finally {
      setLoading(false);
    }
  };

  const handleDelete = async (id: number) => {
    if (confirm("¿Seguro que desea eliminar este registro?")) {
      setLoading(true);
      await service.eliminarMaterial(id);
      await cargarMateriales();
    }
  };

  const handleUpdate = async () => {
    if (!selectedItem) return;
    try {
      setLoading(true);
      // Le pasamos el ID, los datos actualizados y la posible nueva foto (selectedItem.file)
      await service.editarMaterial(selectedItem.Id, {
        Title: selectedItem.Title,
        Categoria: selectedItem.Categoria,
        StockActual: selectedItem.StockActual,
        StockMinimo: selectedItem.StockMinimo
      }, selectedItem.file || null);
      
      setIsPanelOpen(false);
      await cargarMateriales();
    } catch (e) {
      console.error("Error al actualizar", e);
      alert("Error al actualizar el material.");
    } finally {
      setLoading(false);
    }
  };

  const columns: IColumn[] = [
    {
      key: "col0",
      name: "Foto",
      minWidth: 50,
      maxWidth: 50,
      onRender: (item) => {
        let fotoUrl = null;
        if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
            const adjunto = item.AttachmentFiles[0];
            fotoUrl = adjunto.ServerRelativeUrl;
            
            // Reparación de ruta para testeo en entorno local (localhost)
            if (fotoUrl && (window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1")) {
                const tenantUrl = props.context.pageContext.web.absoluteUrl.replace(props.context.pageContext.web.serverRelativeUrl, "");
                fotoUrl = encodeURI(tenantUrl + fotoUrl);
            }
        }

        return fotoUrl ? (
          <img src={fotoUrl} alt={item.Title} className={styles.imageThumbnail} onError={(e) => e.currentTarget.style.display = 'none'} />
        ) : (
          <Icon iconName="Photo2" style={{ fontSize: '20px', color: '#c8c6c4' }} />
        );
      },
    },
    {
      key: "col1",
      name: "Material",
      fieldName: "Title",
      minWidth: 150,
      isResizable: true,
      onRender: (item) => <Text>{item.Title}</Text>
    },
    {
      key: "col2",
      name: "Categoría",
      fieldName: "Categoria",
      minWidth: 100,
    },
    {
      key: "col3",
      name: "Stock Actual",
      fieldName: "StockActual",
      minWidth: 80,
      onRender: (item) => (
        <span className={item.StockActual <= item.StockMinimo ? styles.stockCellAlerta : styles.stockCellNormal}>
          {item.StockActual}
        </span>
      )
    },
    {
      key: "col4",
      name: "Mínimo",
      fieldName: "StockMinimo",
      minWidth: 80,
    },
    {
      key: "col5",
      name: "Acciones",
      minWidth: 100,
      onRender: (item) => (
        <Stack horizontal gap={5}>
          <IconButton 
            iconProps={{ iconName: "Edit" }} 
            onClick={() => { setSelectedItem({...item}); setIsPanelOpen(true); }}
            className={styles.actionBtn}
          />
          <IconButton 
            iconProps={{ iconName: "Delete" }} 
            onClick={() => handleDelete(item.Id)}
            className={styles.deleteBtn}
          />
        </Stack>
      ),
    },
  ];

  // Lógica de filtrado combinada (Texto y Categoría)
  const itemsFiltrados = items.filter(i => {
    const coincideTexto = i.Title?.toLowerCase().includes(filterText.toLowerCase());
    const coincideCategoria = filterCategory === "Todas" || i.Categoria === filterCategory;
    return coincideTexto && coincideCategoria;
  });

  // Opciones para el Dropdown del filtro
  const opcionesFiltroCategoria: IDropdownOption[] = [
    { key: "Todas", text: "Todas las categorías" },
    ...categorias
  ];

  return (
    <div className={styles.container}>
      <div className={styles.formCard}>
        <div className={styles.formTitle}>
          <Icon iconName="BoxAdditionSolid" />
          <Text variant="xLarge">Dar de alta nuevo material</Text>
        </div>
        
        <div className={styles.gridForm}>
          {/* FOTO */}
          <Stack>
            <Text variant="small" style={{ fontWeight: 600, paddingBottom: 6 }}>Imagen</Text>
            <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center">
              <div className={styles.imageUploadContainer} onClick={() => document.getElementById('file-nuevo')?.click()} title="Añadir foto">
                {nuevo.ImagenPreview ? <img src={nuevo.ImagenPreview} alt="Preview" /> : <Icon iconName="Camera" />}
              </div>
              <input type="file" accept="image/*" id="file-nuevo" style={{ display: 'none' }} onChange={(e) => handleFileChange(e, 'nuevo')} />
            </Stack>
          </Stack>

          <TextField
            label="Nombre"
            value={nuevo.nombre}
            onChange={(_, v) => setNuevo({...nuevo, nombre: v || ""})}
            required
          />
          <Dropdown
            label="Categoría"
            options={categorias}
            selectedKey={nuevo.cat}
            onChange={(_, o) => setNuevo({...nuevo, cat: o?.key as string})}
          />
          <TextField
            label="Stock"
            type="number"
            value={nuevo.stock.toString()}
            onChange={(_, v) => setNuevo({...nuevo, stock: parseInt(v || "0")})}
          />
          <TextField
            label="Alerta Mín."
            type="number"
            value={nuevo.stockMin.toString()}
            onChange={(_, v) => setNuevo({...nuevo, stockMin: parseInt(v || "0")})}
          />
          <PrimaryButton
            text="Registrar"
            iconProps={{ iconName: "Save" }}
            onClick={handleAdd}
            className={styles.btnAdd}
            style={{ alignSelf: 'flex-end', marginBottom: '2px' }}
          />
        </div>
      </div>

      <div className={styles.searchSection}>
        <Stack horizontal gap={15} verticalAlign="end">
          <Stack.Item grow={1}>
            <SearchBox
              placeholder="Filtrar materiales por nombre..."
              onChange={(_, v) => setFilterText(v || "")}
            />
          </Stack.Item>
          <Stack.Item>
            <Dropdown
              placeholder="Filtrar por categoría"
              options={opcionesFiltroCategoria}
              selectedKey={filterCategory}
              onChange={(_, o) => setFilterCategory(o?.key as string)}
              styles={{ root: { minWidth: 200 } }}
            />
          </Stack.Item>
        </Stack>
      </div>

      {loading ? (
        <Spinner size={SpinnerSize.large} label="Cargando almacén..." />
      ) : (
        <div className={styles.tableContainer}>
          <DetailsList
            items={itemsFiltrados}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        </div>
      )}

      <Panel
        isOpen={isPanelOpen}
        onDismiss={() => setIsPanelOpen(false)}
        headerText="Editar Material"
        type={PanelType.medium}
      >
        {selectedItem && (
          <Stack gap={15} className={styles.panelStack} style={{ marginTop: 20 }}>
            {/* FOTO EDICIÓN */}
            <Stack horizontalAlign="center">
              <div className={styles.imageUploadContainer} style={{ width: '120px', height: '120px' }} onClick={() => document.getElementById('file-edit')?.click()}>
                  {selectedItem.ImagenPreview || selectedItem.AttachmentFiles?.[0]?.ServerRelativeUrl ? (
                      <img src={selectedItem.ImagenPreview || selectedItem.AttachmentFiles[0].ServerRelativeUrl} alt="Material" />
                  ) : (
                      <Icon iconName="Camera" style={{ fontSize: '30px' }} />
                  )}
              </div>
              <input type="file" accept="image/*" id="file-edit" style={{ display: 'none' }} onChange={(e) => handleFileChange(e, 'edit')} />
              <DefaultButton text="Cambiar Foto" onClick={() => document.getElementById('file-edit')?.click()} style={{ marginTop: '10px' }} />
            </Stack>

            <TextField 
              label="Nombre" 
              value={selectedItem.Title} 
              onChange={(_, v) => setSelectedItem({...selectedItem, Title: v || ""})}
            />
            <Dropdown 
              label="Categoría" 
              options={categorias} 
              selectedKey={selectedItem.Categoria} 
              onChange={(_, o) => setSelectedItem({...selectedItem, Categoria: o?.key as string})}
            />
            <TextField 
              label="Stock Actual" 
              type="number" 
              value={selectedItem.StockActual?.toString()} 
              onChange={(_, v) => setSelectedItem({...selectedItem, StockActual: parseInt(v || "0")})}
            />
            <TextField 
              label="Stock Mínimo" 
              type="number" 
              value={selectedItem.StockMinimo?.toString()} 
              onChange={(_, v) => setSelectedItem({...selectedItem, StockMinimo: parseInt(v || "0")})}
            />
            
            <Separator />
            
            <Stack horizontal gap={10}>
              <PrimaryButton text="Guardar Cambios" onClick={handleUpdate} />
              <DefaultButton text="Cancelar" onClick={() => setIsPanelOpen(false)} />
            </Stack>
          </Stack>
        )}
      </Panel>
    </div>
  );
};

export default ListaMateriales;