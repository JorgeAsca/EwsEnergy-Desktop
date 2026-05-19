import { IPersonal } from "./IPersonal";

export interface IAsignacion {
    Id?: number;
    ObraId: number;
    PersonalId: number;
    Personal?: IPersonal;
    FechaInicio: Date;
    FechaFinPrevista: Date;
    EstadoProgreso: number;
    FechaFinReal?: Date;
    Comentarios?: string;
}