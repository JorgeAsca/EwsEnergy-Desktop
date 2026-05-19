import { StockService } from './StockService';
import { ProjectService } from './ProjectService';

export class LogicService {
    private _stockService: StockService;
    private _projectService: ProjectService;

    constructor(context: any) {
        this._stockService = new StockService(context);
        this._projectService = new ProjectService(context);
    }

    /**
     * Esta función automatiza el descuento de stock cuando se aprueba una obra
     */
    public async procesarSalidaDeMaterial(obraId: number, materialId: number, cantidad: number): Promise<void> {
        // 1. Obtenemos el material actual para ver cuánto stock hay
        const inventario = await this._stockService.getInventario();
        const material = inventario.find(m => m.Id === materialId);

        if (material) {
            const nuevoStock = material.StockActual - cantidad;

            // 2. Si el stock es insuficiente, lanzamos un error (regla de negocio)
            if (nuevoStock < 0) {
                throw new Error("No hay stock suficiente en el almacén para esta obra.");
            }

            // 3. Actualizamos el stock en SharePoint
            await this._stockService.actualizarStock(materialId, nuevoStock);

            // 4. Cambiamos el estado de la obra a 'STOCK ALMACEN'
            await this._projectService.actualizarEstado(obraId, 'STOCK ALMACEN');
            
            console.log(`Proceso completado: Se descontaron ${cantidad} unidades de ${material.Title}`);
        }
    }

    /**
     * Calcula la desviación de materiales (Presupuestado vs Real)
     */
    public calcularDesviacion(presupuestado: number, real: number): number {
        // Retorna el porcentaje de desviación
        if (presupuestado === 0) return 0;
        return ((real - presupuestado) / presupuestado) * 100;
    }
}