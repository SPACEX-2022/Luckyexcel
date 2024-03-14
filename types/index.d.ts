export namespace Luckyexcel {
    /**
     * excel file 转换为 luckysheet json 数据
     */
    function transformExcelToLucky(): void;

    /**
     * excel url 在线地址转换为 luckysheet json 数据
     */
    function transformExcelToLuckyByUrl(): void;

    /**
     * luckysheet json 数据转换为 excel arraybuffer
     */
    function transformLuckyToExcel(): void;
}