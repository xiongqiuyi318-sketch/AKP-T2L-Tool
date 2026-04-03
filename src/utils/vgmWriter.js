import ExcelJS from 'exceljs';

/**
 * Parse VGM file and extract CTN NO., SEAL NO., DESCRIPTION.
 * Falls back to fixed K/L columns for CTN/SEAL if header lookup fails.
 * @param {File} file
 * @returns {Promise<{records: Array<{ctnNo: string, sealNo: string, description: string}>}>}
 */
export async function parseVgmFile(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('VGM文件中没有可用工作表');
  }

  let headerRowIndex = -1;
  for (let rowNum = 1; rowNum <= Math.min(20, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    const rowTextParts = [];
    row.eachCell({ includeEmpty: true }, (cell) => {
      rowTextParts.push(String(cell.value || '').toUpperCase());
    });
    const rowText = rowTextParts.join('|');

    if (
      rowText.includes('CTN') ||
      rowText.includes('CONTAINER') ||
      rowText.includes('SEAL') ||
      rowText.includes('DESCRIPTION')
    ) {
      headerRowIndex = rowNum;
      break;
    }
  }

  if (headerRowIndex === -1) {
    throw new Error('未找到VGM表头行（包含CTN/SEAL/DESCRIPTION）');
  }

  const headerRow = worksheet.getRow(headerRowIndex);
  let ctnIndexCol = -1;
  let ctnCol = -1;
  let sealCol = -1;
  let descriptionCol = -1;

  for (let colNum = 1; colNum <= worksheet.columnCount; colNum++) {
    const headerValue = String(headerRow.getCell(colNum).value || '').toUpperCase().trim();

    if (
      ctnIndexCol === -1 &&
      (headerValue === 'CTN NO.' || headerValue === 'CTN NO' || headerValue === 'CTN')
    ) {
      ctnIndexCol = colNum;
    }

    // Container number column should be the real container ID (e.g. MSMU2289455),
    // not the left-side ctn index (1,2,3...). Prefer CNTR/CONTAINER labels.
    if (
      ctnCol === -1 &&
      (headerValue.includes('CNTR') || headerValue.includes('CONTAINER'))
    ) {
      ctnCol = colNum;
    }

    if (
      sealCol === -1 &&
      (headerValue.includes('SEAL NO') || headerValue === 'SEAL' || headerValue.includes('SEAL'))
    ) {
      sealCol = colNum;
    }

    if (
      descriptionCol === -1 &&
      (headerValue.includes('DESCRIPTION') || headerValue.includes('PRODUCT DESCRIPTION'))
    ) {
      descriptionCol = colNum;
    }
  }

  // If still not found, allow CTN NO. only in right-side area (typically K/L block).
  if (ctnCol === -1) {
    for (let colNum = 1; colNum <= worksheet.columnCount; colNum++) {
      const headerValue = String(headerRow.getCell(colNum).value || '').toUpperCase().trim();
      if (headerValue.includes('CTN NO') && colNum >= 8) {
        ctnCol = colNum;
        break;
      }
    }
  }

  // Fallback to fixed columns from the provided sample: K/L.
  if (ctnCol === -1) ctnCol = 11;
  if (sealCol === -1) sealCol = 12;

  const records = [];
  for (let rowNum = headerRowIndex + 1; rowNum <= worksheet.rowCount; rowNum++) {
    const row = worksheet.getRow(rowNum);
    const ctnIndex = ctnIndexCol !== -1 ? String(row.getCell(ctnIndexCol).value || '').trim() : '';
    const ctnNo = String(row.getCell(ctnCol).value || '').trim();
    const sealNo = String(row.getCell(sealCol).value || '').trim();

    if (!ctnNo && !sealNo) continue;

    let description = 'MARBLE BLOCKS';
    if (descriptionCol !== -1) {
      const sourceDescription = String(row.getCell(descriptionCol).value || '').trim();
      if (sourceDescription) {
        description = sourceDescription;
      }
    }

    records.push({
      ctnIndex,
      ctnNo,
      sealNo,
      description
    });
  }

  if (records.length === 0) {
    throw new Error('VGM中未读取到有效的CTN/SEAL数据');
  }

  return { records };
}

/**
 * Build VGM lookup by container index.
 * Priority: exact index match; fallback by appearance order.
 * @param {Array<{ctnNo: string, matchedStones: Array}>} containersWithData
 * @param {{records: Array<{ctnIndex?: string, ctnNo: string, sealNo: string, description: string}>}} vgmData
 * @returns {Map<string, {ctnNo: string, sealNo: string, description: string} | null>}
 */
function buildVgmMap(containersWithData, vgmData) {
  const byIndex = new Map();
  const noIndexRecords = [];

  vgmData.records.forEach((record) => {
    const indexKey = String(record.ctnIndex || '').trim();
    if (indexKey) {
      byIndex.set(indexKey, record);
    } else {
      noIndexRecords.push(record);
    }
  });

  const mapped = new Map();
  let fallbackPointer = 0;
  containersWithData.forEach((container) => {
    const key = String(container.ctnNo || '').trim();
    if (byIndex.has(key)) {
      mapped.set(key, byIndex.get(key));
    } else {
      mapped.set(key, noIndexRecords[fallbackPointer] || null);
      fallbackPointer += 1;
    }
  });

  return mapped;
}

/**
 * Generate PL WITH CTN NO. based on packing-list rows and enrich with VGM columns.
 * @param {Array<{ctnNo: string, matchedStones: Array}>} containersWithData
 * @param {{records: Array<{ctnNo: string, sealNo: string, description: string}>}} vgmData
 * @returns {Promise<ExcelJS.Workbook>}
 */
export async function generatePlWithCtnNoFromPacking(containersWithData, vgmData) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('PL WITH CTN NO.');

  worksheet.columns = [
    { header: 'CTN', key: 'ctn', width: 10 },
    { header: 'Block No.', key: 'blockNo', width: 15 },
    { header: 'CATE', key: 'cate', width: 10 },
    { header: 'CNTR NO.', key: 'cntrNo', width: 18 },
    { header: 'SEAL NO.', key: 'sealNo', width: 14 },
    { header: 'DESCRIPTION', key: 'description', width: 20 },
    { header: 'BRUTTO DUZINA', key: 'bruttoDuzina', width: 15 },
    { header: 'BRUTTO SIRINA', key: 'bruttoSirina', width: 15 },
    { header: 'BRUTTO VISINA', key: 'bruttoVisina', width: 15 },
    { header: 'BRUTTO M3', key: 'bruttoM3', width: 12 },
    { header: 'NETTO DUZINA', key: 'nettoDuzina', width: 15 },
    { header: 'NETTO SIRINA', key: 'nettoSirina', width: 15 },
    { header: 'NETTO VISINA', key: 'nettoVisina', width: 15 },
    { header: 'NETTO M3', key: 'nettoM3', width: 12 },
    { header: 'WGT.(Tons)', key: 'wgt', width: 12 },
    { header: 'TOTAL WGT.', key: 'totalWgt', width: 12 },
    { header: 'KUPAC', key: 'kupac', width: 15 }
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F4E78' }
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 22;

  const vgmMap = buildVgmMap(containersWithData, vgmData);
  let currentRow = 2;

  containersWithData.forEach((container, containerIndex) => {
    const stones = container.matchedStones || [];
    const vgmInfo = vgmMap.get(String(container.ctnNo).trim());
    const containerTotalWgt = stones.reduce((sum, stone) => sum + (Number(stone.wgt) || 0), 0);
    const groupStartRow = currentRow;

    stones.forEach((stone, stoneIndex) => {
      const row = worksheet.getRow(currentRow);
      const firstStoneInContainer = stoneIndex === 0;

      row.getCell('ctn').value = firstStoneInContainer ? container.ctnNo : '';
      row.getCell('cntrNo').value = firstStoneInContainer ? (vgmInfo?.ctnNo || '') : '';
      row.getCell('sealNo').value = firstStoneInContainer ? (vgmInfo?.sealNo || '') : '';
      row.getCell('description').value = firstStoneInContainer ? (vgmInfo?.description || 'MARBLE BLOCKS') : '';

      row.getCell('blockNo').value = stone.fullBlkNo || `${stone.blkNo || ''}${stone.blkNoYear || ''}`;
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
      row.getCell('totalWgt').value = firstStoneInContainer ? containerTotalWgt : '';
      row.getCell('kupac').value = stone.kupac || '';

      row.getCell('bruttoDuzina').numFmt = '0.00';
      row.getCell('bruttoSirina').numFmt = '0.00';
      row.getCell('bruttoVisina').numFmt = '0.00';
      row.getCell('bruttoM3').numFmt = '0.00';
      row.getCell('nettoDuzina').numFmt = '0.00';
      row.getCell('nettoSirina').numFmt = '0.00';
      row.getCell('nettoVisina').numFmt = '0.00';
      row.getCell('nettoM3').numFmt = '0.00';
      row.getCell('wgt').numFmt = '0.00';
      row.getCell('totalWgt').numFmt = '0.00';

      row.alignment = { vertical: 'middle', horizontal: 'center' };
      row.height = 18;

      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      currentRow += 1;
    });

    const groupEndRow = currentRow - 1;
    if (stones.length > 1 && groupEndRow >= groupStartRow) {
      // Merge CTN (A), CNTR NO. (D), SEAL NO. (E), DESCRIPTION (F), TOTAL WGT. (P) by container group.
      worksheet.mergeCells(`A${groupStartRow}:A${groupEndRow}`);
      worksheet.mergeCells(`D${groupStartRow}:D${groupEndRow}`);
      worksheet.mergeCells(`E${groupStartRow}:E${groupEndRow}`);
      worksheet.mergeCells(`F${groupStartRow}:F${groupEndRow}`);
      worksheet.mergeCells(`P${groupStartRow}:P${groupEndRow}`);
      worksheet.getCell(`A${groupStartRow}`).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getCell(`D${groupStartRow}`).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getCell(`E${groupStartRow}`).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getCell(`F${groupStartRow}`).alignment = { vertical: 'middle', horizontal: 'center' };
      worksheet.getCell(`P${groupStartRow}`).alignment = { vertical: 'middle', horizontal: 'center' };
    }

    if (containerIndex < containersWithData.length - 1) {
      currentRow += 1; // Keep a blank row between containers.
    }
  });

  return workbook;
}

/**
 * Download generated workbook.
 * @param {ExcelJS.Workbook} workbook
 */
export async function downloadPlWithCtnNo(workbook, customFilename) {
  const filename = customFilename || 'PL WITH CTN NO.xlsx';
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
