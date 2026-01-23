import ExcelJS from 'exceljs';
import { formatB10Text, calculateTotalValue } from './dataProcessor.js';

/**
 * 填充单个T2L工作表
 * @param {ExcelJS.Worksheet} templateWorksheet - T2L模板工作表
 * @param {Array} matchedStones - 匹配的石头数据
 * @param {number} t2lNumber - T2L编号
 * @param {number} year - 年份
 * @param {number} containerIndex - 柜序号（从1开始）
 * @returns {ExcelJS.Worksheet} - 填充后的工作表
 */
export function fillT2LSheet(templateWorksheet, matchedStones, t2lNumber, year, containerIndex) {
  // 克隆模板工作表（ExcelJS会自动保留所有格式）
  const worksheet = templateWorksheet.model;
  
  // 填充B10单元格
  const b10Text = formatB10Text(t2lNumber, year, containerIndex);
  const b10Cell = templateWorksheet.getCell('B10');
  b10Cell.value = b10Text;
  // 确保自动换行、居中、加粗
  b10Cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
  b10Cell.font = { bold: true };
  
  // 填充石头数据 (C13:D20, E13:G20, I13:I20, J13:J20)
  matchedStones.forEach((stone, index) => {
    if (index >= 8) return; // 最多8颗石头
    
    const rowNum = 13 + index; // 从第13行开始
    
    // C13:D20 - Block nr. (数字部分 + 年份部分)
    templateWorksheet.getCell(`C${rowNum}`).value = stone.blkNo;
    templateWorksheet.getCell(`D${rowNum}`).value = stone.blkNoYear;
    
    // E13:G20 - Dimension (netto的长、宽、高)
    templateWorksheet.getCell(`E${rowNum}`).value = parseFloat(stone.duzina) || 0;
    templateWorksheet.getCell(`F${rowNum}`).value = parseFloat(stone.sirina) || 0;
    templateWorksheet.getCell(`G${rowNum}`).value = parseFloat(stone.visina) || 0;
    
    // I13:I20 - Category (CATE)
    templateWorksheet.getCell(`I${rowNum}`).value = stone.cate;
    
    // J13:J20 - Tons (WGT.)
    templateWorksheet.getCell(`J${rowNum}`).value = parseFloat(stone.wgt) || 0;
  });
  
  // 填充J24 - 总价值
  const totalValue = calculateTotalValue(matchedStones);
  const j24Cell = templateWorksheet.getCell('J24');
  j24Cell.value = totalValue;
  j24Cell.numFmt = '#,##0.00 "€"';  // 格式：4,929.00 €
  
  // ========== 应用格式设置 ==========
  
  // G10: 换行 + 顶端对齐 + 加粗
  const g10Cell = templateWorksheet.getCell('G10');
  g10Cell.alignment = { wrapText: true, vertical: 'top' };
  g10Cell.font = { bold: true };
  
  // B11: 居中
  const b11Cell = templateWorksheet.getCell('B11');
  b11Cell.alignment = { horizontal: 'center', vertical: 'middle' };
  
  // C12, E12, H12, I12: 居中 + 加粗
  ['C12', 'E12', 'H12', 'I12'].forEach(cellAddr => {
    const cell = templateWorksheet.getCell(cellAddr);
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.font = { bold: true };
  });
  
  // B24, B25, B26, B27, F27, G27: 加粗
  ['B24', 'B25', 'B26', 'B27', 'F27', 'G27'].forEach(cellAddr => {
    const cell = templateWorksheet.getCell(cellAddr);
    cell.font = { bold: true };
  });
  
  // B29: 换行 + 顶端对齐 + 垂直居中
  const b29Cell = templateWorksheet.getCell('B29');
  b29Cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'left' };
  
  // 设置行高
  templateWorksheet.getRow(21).height = 28;  // B21行为28磅
  templateWorksheet.getRow(28).height = 24;  // A28行为24磅
  templateWorksheet.getRow(29).height = 37;  // A29行为37磅
  
  // 设置打印选项
  templateWorksheet.pageSetup = {
    paperSize: 9,  // A4
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    printArea: 'A1:I29'
  };
  
  // 设置B10:I29的外侧框线
  const borderStyle = { style: 'thin', color: { argb: 'FF000000' } };
  
  // 顶部边框 (B10:I10)
  for (let col = 2; col <= 9; col++) { // B=2, I=9
    const cell = templateWorksheet.getCell(10, col);
    if (!cell.border) cell.border = {};
    cell.border.top = borderStyle;
  }
  
  // 底部边框 (B29:I29)
  for (let col = 2; col <= 9; col++) {
    const cell = templateWorksheet.getCell(29, col);
    if (!cell.border) cell.border = {};
    cell.border.bottom = borderStyle;
  }
  
  // 左侧边框 (B10:B29)
  for (let row = 10; row <= 29; row++) {
    const cell = templateWorksheet.getCell(row, 2); // B列
    if (!cell.border) cell.border = {};
    cell.border.left = borderStyle;
  }
  
  // 右侧边框 (I10:I29)
  for (let row = 10; row <= 29; row++) {
    const cell = templateWorksheet.getCell(row, 9); // I列
    if (!cell.border) cell.border = {};
    cell.border.right = borderStyle;
  }
  
  return templateWorksheet;
}

/**
 * 生成包含多个工作表的Excel文件
 * @param {ExcelJS.Workbook} templateWorkbook - 模板workbook
 * @param {ExcelJS.Worksheet} templateWorksheet - T2L模板工作表
 * @param {Array} containers - 柜数组 [{ctnNo, matchedStones}, ...]
 * @param {number} startNumber - 起始T2L编号
 * @param {number} year - 年份
 * @returns {ExcelJS.Workbook} - 完整的workbook对象
 */
export function buildWorkbookWithSheets(templateWorkbook, templateWorksheet, containers, startNumber, year) {
  const workbook = new ExcelJS.Workbook();
  
  containers.forEach((container, index) => {
    const t2lNumber = startNumber + index;
    
    // 复制模板工作表（ExcelJS会保留所有格式）
    const newWorksheet = workbook.addWorksheet(container.ctnNo);
    
    // 复制模板的所有内容和格式
    templateWorksheet.eachRow((row, rowNumber) => {
      const newRow = newWorksheet.getRow(rowNumber);
      newRow.height = row.height;
      
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        
        // 复制值
        newCell.value = cell.value;
        
        // 复制样式
        if (cell.style) {
          newCell.style = JSON.parse(JSON.stringify(cell.style));
        }
      });
    });
    
    // 复制列宽
    templateWorksheet.columns.forEach((col, index) => {
      if (col.width) {
        newWorksheet.getColumn(index + 1).width = col.width;
      }
    });
    
    // 复制合并单元格
    if (templateWorksheet.model.merges) {
      templateWorksheet.model.merges.forEach(merge => {
        newWorksheet.mergeCells(merge);
      });
    }
    
    // 填充数据（containerIndex从1开始）
    const containerIndex = index + 1;
    fillT2LSheet(newWorksheet, container.matchedStones, t2lNumber, year, containerIndex);
  });
  
  return workbook;
}

/**
 * 下载Excel文件
 * @param {ExcelJS.Workbook} workbook - workbook对象
 * @param {string} filename - 文件名
 */
export async function downloadExcel(workbook, filename = 'T2L_Output.xlsx') {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  });
  
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
