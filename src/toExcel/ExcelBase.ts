import {IExcel} from "./IExcel";
import {Workbook} from "exceljs/index.d";

export class ExcelBase implements IExcel {
    workbook: Workbook;
}