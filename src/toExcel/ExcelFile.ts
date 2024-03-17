
import Excel from 'exceljs';
import { LuckyFileBase } from "../ToLuckySheet/LuckyBase";
import { ILuckyFile, ILuckyFileInfo, IluckySheet } from "../ToLuckySheet/ILuck";
import { IdownloadfileList } from "../ICommon";
import {getColumnWidthExcel, getRowHeightExcel} from "../common/method";
import {Worksheet} from "./Worksheet";

export class ExcelFile implements ILuckyFile{
    
    info:ILuckyFileInfo;
    sheets:IluckySheet[];
    
    constructor(luckyFile:ILuckyFile){
        // super();
        this.info = luckyFile.info;
        this.sheets = luckyFile.sheets;
    }

    async Parse() {
        const sheets = this.sheets;
        // todo: transform json to xml string
        
        // relsFile toRels()
        // workBookFile toWorkBook()
        // stylesFile toStyles()
        // workbookRels toWorkBookRels()
        // contentTypesFile toContentType()
        // worksheetFilePath toWorkSheets()
        // 1.创建工作簿，可以为工作簿添加属性
        const workbook = new Excel.Workbook()
        // 2.创建表格，第二个参数可以配置创建什么样的工作表
        //@ts-ignore
        sheets.every(function (table) {
            if (table.data.length === 0) return true
            const worksheet = new Worksheet(workbook, table)
            return true
        })
        // 4.写入 buffer
        const buffer = await workbook.xlsx.writeBuffer()

        // console.log(sheets);
        return buffer;
    }
}