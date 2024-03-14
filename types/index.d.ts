import {IuploadfileList} from "../src/ICommon";
import {ILuckyJson} from "../src/ToLuckySheet/ILuck";

export namespace LuckyExcel {
    /**
     * excel file 转换为 luckysheet json 数据
     */
    function transformExcelToLucky(excelFile:File, callBack?:(files:IuploadfileList, fs?:string)=>void): void;

    /**
     * excel url 在线地址转换为 luckysheet json 数据
     */
    function transformExcelToLuckyByUrl(url:string, name:string, callBack?:(files:IuploadfileList, fs?:string)=>void): void;

    /**
     * luckysheet json 数据转换为 excel arraybuffer
     */
    function transformLuckyToExcel(luckyJson: ILuckyJson, callBack?: (content: ArrayBuffer, title: string) => void): void;
}