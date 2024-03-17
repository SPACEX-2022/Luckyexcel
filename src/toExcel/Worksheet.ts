
//转换颜色
import {getColumnWidthExcel, getRowHeightExcel} from "../common/method";
import {IExcel} from "./IExcel";
import {ExcelBase} from "./ExcelBase";
import {Workbook} from "exceljs/index.d";
import {IluckySheet} from "../ToLuckySheet/ILuck";

const rgb2hex = (rgb: any) => {
    if (rgb.charAt(0) == '#') {
        return rgb
    }

    var ds = rgb.split(/\D+/)
    var decimal = Number(ds[1]) * 65536 + Number(ds[2]) * 256 + Number(ds[3])
    return "#" + zero_fill_hex(decimal, 6)

    function zero_fill_hex(num: any, digits: any) {
        var s = num.toString(16)
        while (s.length < digits)
            s = "0" + s
        return s
    }
}
const fillConvert = (bg: any) => {
    if (!bg) {
        return {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '#ffffff'.replace('#', '') }
        }
    }
    bg = bg.indexOf('rgb') > -1 ? rgb2hex(bg) : bg;
    let fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bg.replace('#', '') }
    }
    return fill
}
const setMerge = (luckyMerge = {}, worksheet: any) => {
    const mergearr = Object.values(luckyMerge)
    mergearr.forEach(function (elem) { // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
        // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
        //@ts-ignore
        worksheet.mergeCells(elem.r + 1, elem.c + 1, elem.r + elem.rs, elem.c + elem.cs);
    })
}
const setBorder = (luckyBorderInfo: any, worksheet: any) => {
    if (!Array.isArray(luckyBorderInfo)) return
    luckyBorderInfo.forEach((elem) => {
        var val = elem;
        let border: any = {}
        const luckyToExcel: any = {
            type: {
                'border-top': 'top',
                'border-right': 'right',
                'border-bottom': 'bottom',
                'border-left': 'left'
            },
            style: {
                0: 'none',
                1: 'thin',
                2: 'hair',
                3: 'dotted',
                4: 'dashDot',
                5: 'dashDot',
                6: 'dashDotDot',
                7: 'double',
                8: 'medium',
                9: 'mediumDashed',
                10: 'mediumDashDot',
                11: 'mediumDashDotDot',
                12: 'slantDashDot',
                13: 'thick'
            }
        }

        if (val.rangeType === 'range') {
            let color = (val.color.replace('#', '')).toUpperCase()
            if (val.borderType === 'border-all') {
                border['top'] = { style: luckyToExcel.style[val.style], color: { argb: color } }
                border['right'] = { style: luckyToExcel.style[val.style], color: { argb: color } }
                border['bottom'] = { style: luckyToExcel.style[val.style], color: { argb: color } }
                border['left'] = { style: luckyToExcel.style[val.style], color: { argb: color } }
            } else {
                border[luckyToExcel.type[val.borderType]] = { style: luckyToExcel.style[val.style], color: { argb: color } }
            }
            val.range.forEach((item: any) => {
                for (let r = item.row[0]; r < item.row[1] + 1; r++) {
                    for (let c = item.column[0]; c < item.column[1] + 1; c++) {
                        worksheet.getCell(r + 1, c + 1).border = border
                    }
                }
            })
        } else if (val.rangeType === 'cell') {
            if (val.value.t) border['top'] = { style: luckyToExcel.style[val.value.t.style], color: { argb: (val.value.t.color.replace('#', '')).toUpperCase() } }
            if (val.value.r) border['right'] = { style: luckyToExcel.style[val.value.r.style], color: { argb: (val.value.r.color.replace('#', '')).toUpperCase() } }
            if (val.value.b) border['bottom'] = { style: luckyToExcel.style[val.value.b.style], color: { argb: (val.value.b.color.replace('#', '')).toUpperCase() } }
            if (val.value.l) border['left'] = { style: luckyToExcel.style[val.value.l.style], color: { argb: (val.value.l.color.replace('#', '')).toUpperCase() } }
            worksheet.getCell(val.value.row_index + 1, val.value.col_index + 1).border = border
        }
    })
}

const setStyleAndValue = (cellArr: any, worksheet: any) => {
    if (!Array.isArray(cellArr)) return;

    cellArr.forEach(function (row, rowid) {
        const dbrow = worksheet.getRow(rowid + 1);
        //设置单元格行高,默认除以1.3倍
        //@ts-ignore
        dbrow.height = getRowHeightExcel(luckysheet.getRowHeight([rowid])[rowid]);
        row.every(function (cell: any, columnid: any) {
            if (!cell) return true;
            // 根据前三行的元素设置单元格宽度 有可能第一行为空
            if (rowid < 3) {
                const dobCol = worksheet.getColumn(columnid + 1);
                //设置单元格列宽除以8
                //@ts-ignore
                dobCol.width = getColumnWidthExcel(luckysheet.getColumnWidth([columnid])[columnid]);
            }
            let fill = fillConvert(cell.bg);
            let font = fontConvert(cell.ff, cell.fc, cell.bl, cell.it, cell.fs, cell.cl, cell.ul);
            let alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr);
            let value;

            var v = '';
            if (cell.ct && cell.ct.t == 'inlineStr') {
                var s = cell.ct.s;
                s.forEach(function (val: any, num: any) {
                    v += val.v;
                })
            } else {
                v = cell.v;
            }
            if (cell.f) {
                value = { formula: cell.f, result: v };
            } else {
                value = v;
            }
            let target = worksheet.getCell(rowid + 1, columnid + 1);
            target.fill = fill;
            target.font = font;
            target.alignment = alignment;
            target.value = value;
            return true;
        })
    })
}
var fontConvert = function (ff = 0, fc = '#000000', bl = 0, it = 0, fs = 10, cl = 0, ul = 0) { // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
    const luckyToExcel = {
        0: '微软雅黑',
        1: '宋体（Song）',
        2: '黑体（ST Heiti）',
        3: '楷体（ST Kaiti）',
        4: '仿宋（ST FangSong）',
        5: '新宋体（ST Song）',
        6: '华文新魏',
        7: '华文行楷',
        8: '华文隶书',
        9: 'Arial',
        10: 'Times New Roman ',
        11: 'Tahoma ',
        12: 'Verdana',
        num2bl: function (num: number) {
            return num === 0 ? false : true
        }
    }

    let font = {
        name: ff,
        family: 1,
        size: fs,
        color: { argb: fc.replace('#', '') },
        bold: luckyToExcel.num2bl(bl),
        italic: luckyToExcel.num2bl(it),
        underline: luckyToExcel.num2bl(ul),
        strike: luckyToExcel.num2bl(cl)
    }

    return font;
}

const alignmentConvert = (vt = 'default', ht = 'default', tb = 'default', tr = 'default') => { // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
    const luckyToExcel: any = {
        vertical: {
            0: 'middle',
            1: 'top',
            2: 'bottom',
            default: 'top'
        },
        horizontal: {
            0: 'center',
            1: 'left',
            2: 'right',
            default: 'left'
        },
        wrapText: {
            0: false,
            1: false,
            2: true,
            default: false
        },
        textRotation: {
            0: 0,
            1: 45,
            2: -45,
            3: 'vertical',
            4: 90,
            5: -90,
            default: 0
        }
    }
    let alignment = {
        vertical: luckyToExcel.vertical[vt],
        horizontal: luckyToExcel.horizontal[ht],
        wrapText: luckyToExcel.wrapText[tb],
        textRotation: luckyToExcel.textRotation[tr]
    }
    return alignment;
}


export class Worksheet extends ExcelBase {
    constructor(workbook: Workbook, table: IluckySheet) {
        super();
        this.workbook = workbook;
        const worksheet = this.workbook.addWorksheet(table.name, {
            views: [{
                    showGridLines: table.showGridLines !== '0',
                }],
            properties: {
                defaultColWidth: getColumnWidthExcel(table.defaultColWidth),
                defaultRowHeight: getRowHeightExcel(table.defaultRowHeight)
            }
        })

        // 3.设置单元格合并,设置单元格边框,设置单元格样式,设置值
        setStyleAndValue(table.data, worksheet)
        setMerge(table.config.merge, worksheet)
        setBorder(table.config.borderInfo, worksheet)

        return worksheet;
    }
}