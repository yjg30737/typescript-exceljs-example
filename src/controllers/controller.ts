import { Request, Response } from "express";
import excel from "exceljs";
import * as service from '../services/service';

export async function getAll(req: Request, res: Response): Promise<void> {
    try {
        let result = await service.getAll();
        res.status(200).send(result);
    } catch (e) {
        res.sendStatus(500);
    }
}

export async function getById(req: Request, res: Response): Promise<void> {
    try {
        let id: number = Number(req.query['id']);
        let result = await service.getById(id);
        res.status(200).send(result);
    } catch (e) {
        res.sendStatus(500);
    }
}

export async function getExcel(req: Request, res: Response): Promise<void> {
    try {
        const dataHeaderObj = [
            { header: '아이디', key: 'id'},
            { header: '이름', key: 'name'},
            { header: '나이', key: 'age'},
            { header: '직업', key: 'job'},
        ];
        const dataValueObj = await service.getAll();
        const dataResultObj = JSON.parse(JSON.stringify(dataValueObj));

        const workbook = new excel.Workbook();
        const sheet = workbook.addWorksheet('My Sheet');

        const dataFirstCol = 1; 
        const dataFirstRow = 1;
        const dataLastCol = dataHeaderObj.length+dataFirstCol-1;
        const dataLastRow = dataFirstRow+dataResultObj.length;

        sheet.columns = dataHeaderObj;
        sheet.addRows(dataResultObj);
        
        function colName(num: number) {
            let letters = ''
            while (num >= 0) {
                letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[num % 26] + letters
                num = Math.floor(num / 26) - 1
            }
            return letters
        }
 
        function getCellStr(col: number, row: number, colOffset: number = 0, rowOffset: number = 0) {
            return `${colName(col+colOffset)}${row+rowOffset}`
        }
        
        function getHeaderList(startCol: number, startRow: number, endCol: number) {
            let startColIdx = startCol-1
            let headerList = []
            
            for(let n = startColIdx; n < endCol; n++) {
                headerList.push(getCellStr(n, startRow));   
            }
            return headerList;
        }

        const dataHeaderCellList = getHeaderList(dataFirstCol, dataFirstRow, dataLastCol); 

        const alignStyle: Partial<excel.Alignment> = { vertical: 'middle', horizontal: 'center' };

        const borderStyle: Partial<excel.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }

        const font = { name: 'Arial', size: 16, bold: true };

        const color: any = { type: 'gradient',
            gradient: 'angle',
            degree: 0,
            stops: [
                {position:0, color:{argb:'FF0000FF'}},
                {position:0.5, color:{argb:'FFFFFFFF'}},
                {position:1, color:{argb:'FF0000FF'}}
        ]};

        dataHeaderCellList.forEach((cell: string) => {
            sheet.getCell(cell).alignment = alignStyle;
            sheet.getCell(cell).border = borderStyle;
            sheet.getCell(cell).font = font;
            sheet.getCell(cell).fill = color;
        })
        
        const createOuterBorder = (worksheet: any, start = {row: 1, col: 1}, end = {row: 1, col: 1}, borderWidth = 'thin') => {
            const borderStyle = {
                style: borderWidth
            };
            for (let i = start.row; i <= end.row; i++) {
                const leftBorderCell = worksheet.getCell(i, start.col);
                const rightBorderCell = worksheet.getCell(i, end.col);
                leftBorderCell.border = {
                    ...leftBorderCell.border,
                    left: borderStyle
                };
                rightBorderCell.border = {
                    ...rightBorderCell.border,
                    right: borderStyle
                };
            }
        
            for (let i = start.col; i <= end.col; i++) {
                const topBorderCell = worksheet.getCell(start.row, i);
                const bottomBorderCell = worksheet.getCell(end.row, i);
                topBorderCell.border = {
                    ...topBorderCell.border,
                    top: borderStyle
                };
                bottomBorderCell.border = {
                    ...bottomBorderCell.border,
                    bottom: borderStyle
                };
            }
        };
        
        
        createOuterBorder(sheet, {row: dataFirstRow, col: dataFirstCol}, {row: dataLastRow, col: dataLastCol});

        
        sheet.columns.forEach(function (column: any, i) {
            let maxLength = 0;
            if(column) {
                column["eachCell"]({ includeEmpty: false }, function (cell: any) {
                    let columnLength = cell.value ? cell.value.toString().length * 2 : 10;
                    if (columnLength > maxLength ) {
                        maxLength = columnLength;
                    }
                });
                column.width = maxLength < 10 ? 10 : maxLength;
            }
        });

        
        
        const totalFirstCol = dataLastCol+2;
        const totalFirstRow = dataFirstRow+3;
        const totalLastCol = totalFirstCol+1;
        const totalLastRow = totalFirstRow+2;

        
        const totalHeaderCellList = getHeaderList(totalFirstCol+1, totalFirstRow, totalLastCol+1);

        
        sheet.getCell(totalHeaderCellList[0]).alignment = alignStyle;
        sheet.getCell(totalHeaderCellList[0]).border = borderStyle;
        sheet.getCell(totalHeaderCellList[0]).font = font;
        sheet.getCell(totalHeaderCellList[0]).fill = color;
        sheet.getCell(totalHeaderCellList[0]).value = '계';
        sheet.mergeCells(`${totalHeaderCellList[0]}:${totalHeaderCellList[totalHeaderCellList.length-1]}`);
        
        sheet.getColumn(colName(totalFirstCol)).width = 15;

        
        const ageCellHeaderStr = getCellStr(totalFirstCol, totalFirstRow, 0, 1);
        const jobCellHeaderStr = getCellStr(totalFirstCol, totalFirstRow, 0, 2);
        sheet.getCell(ageCellHeaderStr).value = '나이';
        sheet.getCell(jobCellHeaderStr).value = '직업 있는 사람';

        
        const ageCellValueStr = getCellStr(totalLastCol, totalFirstRow, 0, 1);
        const jobCellValueStr = getCellStr(totalLastCol, totalFirstRow, 0, 2);
        
        sheet.getCell(ageCellValueStr).value = { formula: `SUM(${sheet.getColumnKey('age').letter}${dataFirstRow+1}:${sheet.getColumnKey('age').letter}${dataLastRow})`, 
                                                 date1904: false };
        
        sheet.getCell(jobCellValueStr).value = { formula: `SUMPRODUCT(--(${sheet.getColumnKey('job').letter}${dataFirstRow+1}:${sheet.getColumnKey('job').letter}${dataLastRow}<>""))`, 
                                                 date1904: false };

        
        createOuterBorder(sheet, {row: totalFirstRow, col: totalFirstCol+1}, {row: totalLastRow, col: totalLastCol+1});

        
        await workbook.xlsx.writeFile('abc.xlsx');

        res.status(200).send({ message: 'good'});
    } catch (e) {
        console.log(e);
        res.sendStatus(500);
    }
}

export async function create(req: Request, res: Response): Promise<void> {
    try {
        await service.create(req.body);
        res.sendStatus(200);
    } catch (e) {
        res.sendStatus(500);
    }
}


export async function update(req: Request, res: Response): Promise<void> {
    try {
        let id: number = Number(req.query['id']);
        await service.update(id, req.body);
        res.sendStatus(200);
    } catch (e) {
        res.sendStatus(500);
    }
}


export async function remove(req: Request, res: Response): Promise<void> {
    try {
        let id: number = Number(req.query['id']);
        await service.remove(id);
        res.sendStatus(200);
    } catch (e) {
        res.sendStatus(500);
    }
}