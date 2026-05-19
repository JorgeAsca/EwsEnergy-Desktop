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
  const [filterCategory, setFilterCategory] = React.useState<string>("Todas"); // <-- NUEVO ESTADO PARA CATEGORÍA
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [selectedItem, setSelectedItem] = React.useState<any>(null);
  
  const [nuevo, setNuevo] = React.useState({
    nombre: "",
    stock: 0,
    stockMin: 0,
    cat: "Consumible",
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

  const handleAdd = async () => {
    if (!nuevo.nombre) return;
    try {
      await service.crearMaterial({
        Title: nuevo.nombre,
        StockActual: nuevo.stock,
        StockMinimo: nuevo.stockMin,
        Categoria: nuevo.cat
      });
      setNuevo({ nombre: "", stock: 0, stockMin: 0, cat: "Consumible", file: null });
      cargarMateriales();
    } catch (e) {
      console.error(e);
    }
  };

  const handleDelete = async (id: number) => {
    if (confirm("¿Seguro que desea eliminar este registro?")) {
      await service.eliminarMaterial(id);
      cargarMateriales();
    }
  };

  const handleUpdate = async () => {
    if (!selectedItem) return;
    try {
      setLoading(true);
      await service.actualizarMaterial(selectedItem.Id, {
        Title: selectedItem.Title,
        Categoria: selectedItem.Categoria,
        StockActual: selectedItem.StockActual,
        StockMinimo: selectedItem.StockMinimo
      });
      setIsPanelOpen(false);
      await cargarMateriales();
    } catch (e) {
      console.error("Error al actualizar", e);
    } finally {
      setLoading(false);
    }
  };

  const columns: IColumn[] = [
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
          <Stack gap={15} className={styles.panelStack}>
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