import ExcelJS from 'exceljs';

/**
 * 生成Packing List Excel文件
 * @param {Array} containers - 柜数组，包含匹配的石头数据
 * @returns {ExcelJS.Workbook} - 生成的workbook
 */
export async function generatePackingList(containers) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Packing List');
  
  // 定义列
  worksheet.columns = [
    { header: 'CTN', key: 'ctn', width: 10 },
    { header: 'Block No.', key: 'blockNo', width: 15 },
    { header: 'CATE', key: 'cate', width: 10 },
    { header: 'BRUTTO DUZINA', key: 'bruttoDuzina', width: 15 },
    { header: 'BRUTTO SIRINA', key: 'bruttoSirina', width: 15 },
    { header: 'BRUTTO VISINA', key: 'bruttoVisina', width: 15 },
    { header: 'BRUTTO M3', key: 'bruttoM3', width: 12 },
    { header: 'NETTO DUZINA', key: 'nettoDuzina', width: 15 },
    { header: 'NETTO SIRINA', key: 'nettoSirina', width: 15 },
    { header: 'NETTO VISINA', key: 'nettoVisina', width: 15 },
    { header: 'NETTO M3', key: 'nettoM3', width: 12 },
    { header: 'WGT.(Tons)', key: 'wgt', width: 12 },
    { header: 'KUPAC', key: 'kupac', width: 15 },
    { header: 'Unit Price', key: 'unitPrice', width: 15 },
    { header: 'Total Price', key: 'totalPrice', width: 15 }
  ];
  
  // 设置表头样式
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' }
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 20;
  
  // 添加数据
  let currentRow = 2;
  
  containers.forEach((container, containerIndex) => {
    const stones = container.matchedStones;
    
    stones.forEach((stone, stoneIndex) => {
      const row = worksheet.getRow(currentRow);
      
      // 只在每组的第一块石头显示CTN号
      row.getCell('ctn').value = stoneIndex === 0 ? container.ctnNo : '';
      row.getCell('blockNo').value = stone.fullBlkNo || `${stone.blkNo}${stone.blkNoYear}`;
      row.getCell('cate').value = stone.cate || '';
      row.getCell('bruttoDuzina').value = stone.bruttoDuzina || 0;
      row.getCell('bruttoSirina').value = stone.bruttoSirina || 0;
      row.getCell('bruttoVisina').value = stone.bruttoVisina || 0;
      row.getCell('bruttoM3').value = stone.bruttoM3 || 0;
      row.getCell('nettoDuzina').value = stone.nettoDuzina || stone.duzina || 0;
      row.getCell('nettoSirina').value = stone.nettoSirina || stone.sirina || 0;
      row.getCell('nettoVisina').value = stone.nettoVisina || stone.visina || 0;
      row.getCell('nettoM3').value = stone.nettoM3 || 0;
      row.getCell('wgt').value = stone.wgt || 0;
      row.getCell('kupac').value = stone.kupac || '';
      row.getCell('unitPrice').value = stone.unitPrice || 0;
      row.getCell('totalPrice').value = stone.totalPrice || 0;
      
      // 设置数字格式
      row.getCell('bruttoDuzina').numFmt = '0.00';
      row.getCell('bruttoSirina').numFmt = '0.00';
      row.getCell('bruttoVisina').numFmt = '0.00';
      row.getCell('bruttoM3').numFmt = '0.00';
      row.getCell('nettoDuzina').numFmt = '0.00';
      row.getCell('nettoSirina').numFmt = '0.00';
      row.getCell('nettoVisina').numFmt = '0.00';
      row.getCell('nettoM3').numFmt = '0.00';
      row.getCell('wgt').numFmt = '0.00';
      row.getCell('unitPrice').numFmt = '#,##0.00 "€"';
      row.getCell('totalPrice').numFmt = '#,##0.00 "€"';
      
      // 设置行样式
      row.alignment = { vertical: 'middle' };
      row.height = 18;
      
      // 添加边框
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      
      currentRow++;
    });
    
    // 在每组之间添加空行（除了最后一组）
    if (containerIndex < containers.length - 1) {
      currentRow++;
    }
  });
  
  // 添加所有单元格边框
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (!cell.border) {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
    });
  });
  
  return workbook;
}

/**
 * 下载Packing List Excel文件
 * @param {ExcelJS.Workbook} workbook - workbook对象
 * @param {number} containerCount - 柜数
 */
export async function downloadPackingList(workbook, containerCount) {
  const filename = `Packing List_${containerCount}.xlsx`;
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
