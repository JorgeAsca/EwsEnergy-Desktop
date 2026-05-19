import { IObra } from "./IObra";
import { IFacepilePersona } from "@fluentui/react";

export interface IObraCard extends IObra {
  clienteNombre: string;
  porcentajeReal: number;
  operarios: IFacepilePersona[];
  jornadasConsumidas: number;
}