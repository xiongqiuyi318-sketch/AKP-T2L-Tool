import ExcelJS from 'exceljs';

/**
 * 解析文件1 - 石头信息
 * @param {File} file - Excel文件
 * @returns {Promise<Object>} - 以BLK NO.为key的石头数据字典
 */
export async function parseStoneInfo(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0];
  
  // 查找表头行
  let headerRowIndex = -1;
  let headers = {};
  
  // 遍历前20行查找表头
  for (let rowNum = 1; rowNum <= Math.min(20, worksheet.rowCount); rowNum++) {
    const row = worksheet.getRow(rowNum);
    
    // 安全地构建行字符串
    const rowStr = [];
    row.eachCell({ includeEmpty: true }, (cell) => {
      rowStr.push(String(cell.value || ''));
    });
    const rowString = rowStr.join('|').toUpperCase();
    
    if (rowString.includes('BLK NO') || rowString.includes('BLOCK NO') || rowString.includes('荒料号')) {
      headerRowIndex = rowNum;
      console.log(`文件1 - 找到表头行: 第${rowNum}行`);
      
      // 建立列索引映射
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const cellStr = String(cell.value || '').toUpperCase();
        console.log(`文件1 - 表头第${colNumber}列: ${cellStr}`);
        
        // 只取第一个匹配的列（避免合并单元格导致的重复）
        if ((cellStr.includes('BLK NO') || cellStr.includes('BLOCK NO') || cellStr.includes('荒料号')) && !headers.blkNo) {
          headers.blkNo = colNumber;
          console.log(`文件1 - BLK NO列确定为: ${colNumber}`);
        } else if (cellStr.includes('CATE') || cellStr.includes('等级')) {
          headers.cate = colNumber;
        } else if (cellStr.includes('NETTO')) {
          headers.nettoStart = colNumber;
        } else if (cellStr.includes('UNIT PRICE') || cellStr.includes('单价')) {
          headers.unitPrice = colNumber;
        }
      });
      
      break; // 找到表头后退出循环
    }
  }
  
  if (headerRowIndex === -1) {
    throw new Error('未找到表头行（需要包含"BLK NO."或"BLOCK NO."字段）');
  }
  
  console.log('文件1 - 表头行索引:', headerRowIndex);
  console.log('文件1 - headers对象:', headers);
  
  // 查找BRUTTO和NETTO区域的尺寸列，以及其他列
  let bruttoDuzina = -1, bruttoSirina = -1, bruttoVisina = -1, bruttoM3 = -1;
  let nettoDuzina = -1, nettoSirina = -1, nettoVisina = -1, nettoM3 = -1;
  let wgt = -1, totalPrice = -1, kupac = -1;
  let bruttoColStart = -1, nettoColStart = -1;
  
  const searchStartRow = Math.max(1, headerRowIndex - 2);
  const searchEndRow = Math.min(headerRowIndex + 2, worksheet.rowCount);
  
  // 第一遍：查找BRUTTO和NETTO的起始列（精确匹配，避免误判）
  for (let rowNum = searchStartRow; rowNum <= searchEndRow; rowNum++) {
    const row = worksheet.getRow(rowNum);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const cellStr = String(cell.value || '').toUpperCase();
      
      // 只在bruttoColStart还未找到时查找BRUTTO
      if (cellStr.includes('BRUTTO') && bruttoColStart === -1) {
        bruttoColStart = colNumber;
        console.log(`文件1 - 找到BRUTTO区域起始列: 第${colNumber}列，表头内容: "${cell.value}"`);
      }
      // 只在nettoColStart还未找到时查找NETTO（排除WGT等干扰）
      if (cellStr.includes('NETTO') && !cellStr.includes('WGT') && nettoColStart === -1) {
        nettoColStart = colNumber;
        console.log(`文件1 - 找到NETTO区域起始列: 第${colNumber}列，表头内容: "${cell.value}"`);
      }
    });
  }
  
  // 第二遍：在BRUTTO和NETTO区域查找具体的列，以及其他列
  for (let rowNum = searchStartRow; rowNum <= searchEndRow; rowNum++) {
    const row = worksheet.getRow(rowNum);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const cellStr = String(cell.value || '').toUpperCase();
      
      // BRUTTO区域的列（在bruttoColStart附近，直到nettoColStart之前）
      const bruttoRangeEnd = nettoColStart !== -1 ? nettoColStart : colNumber + 10;
      if (bruttoColStart !== -1 && colNumber >= bruttoColStart && colNumber < bruttoRangeEnd) {
        if ((cellStr.includes('DUZINA') || cellStr.includes('长')) && bruttoDuzina === -1) {
          bruttoDuzina = colNumber;
          console.log(`文件1 - 找到BRUTTO DUZINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('SIRINA') || cellStr.includes('宽')) && bruttoSirina === -1) {
          bruttoSirina = colNumber;
          console.log(`文件1 - 找到BRUTTO SIRINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('VISINA') || cellStr.includes('高')) && bruttoVisina === -1) {
          bruttoVisina = colNumber;
          console.log(`文件1 - 找到BRUTTO VISINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('M3') || cellStr.includes('立方')) && bruttoM3 === -1) {
          bruttoM3 = colNumber;
          console.log(`文件1 - 找到BRUTTO M3列: 第${colNumber}列`);
        }
      }
      
      // NETTO区域的列（在nettoColStart附近）
      if (nettoColStart !== -1 && colNumber >= nettoColStart && colNumber < nettoColStart + 6) {
        if ((cellStr.includes('DUZINA') || cellStr.includes('长')) && nettoDuzina === -1) {
          nettoDuzina = colNumber;
          console.log(`文件1 - 找到NETTO DUZINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('SIRINA') || cellStr.includes('宽')) && nettoSirina === -1) {
          nettoSirina = colNumber;
          console.log(`文件1 - 找到NETTO SIRINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('VISINA') || cellStr.includes('高')) && nettoVisina === -1) {
          nettoVisina = colNumber;
          console.log(`文件1 - 找到NETTO VISINA列: 第${colNumber}列`);
        } else if ((cellStr.includes('M3') || cellStr.includes('立方')) && nettoM3 === -1) {
          nettoM3 = colNumber;
          console.log(`文件1 - 找到NETTO M3列: 第${colNumber}列`);
        }
      }
      
      // 其他列（不受BRUTTO/NETTO限制，在整个范围查找）
      if ((cellStr.includes('WGT') || cellStr.includes('重量') || cellStr.includes('吨')) && !cellStr.includes('TOTAL') && wgt === -1) {
        wgt = colNumber;
        console.log(`文件1 - 找到WGT列: 第${colNumber}列`);
      }
      
      if ((cellStr.includes('UNIT') && cellStr.includes('PRICE')) || cellStr.includes('单价')) {
        if (!headers.unitPrice) {
          headers.unitPrice = colNumber;
          console.log(`文件1 - 找到UNIT PRICE列: 第${colNumber}列，表头内容: "${cellStr}"`);
        }
      }
      
      if ((cellStr.includes('TOTAL') && cellStr.includes('PRICE')) && totalPrice === -1) {
        totalPrice = colNumber;
        console.log(`文件1 - 找到TOTAL PRICE列: 第${colNumber}列`);
      }
      
      if ((cellStr.includes('KUPAC') || cellStr.includes('客户')) && kupac === -1) {
        kupac = colNumber;
        console.log(`文件1 - 找到KUPAC列: 第${colNumber}列，表头内容: "${cellStr}"`);
      }
    });
  }
  
  // 直接使用找到的BRUTTO/NETTO列
  headers.bruttoDuzina = bruttoDuzina;
  headers.bruttoSirina = bruttoSirina;
  headers.bruttoVisina = bruttoVisina;
  headers.bruttoM3 = bruttoM3;
  headers.nettoDuzina = nettoDuzina;
  headers.nettoSirina = nettoSirina;
  headers.nettoVisina = nettoVisina;
  headers.nettoM3 = nettoM3;
  headers.kupac = kupac;
  
  // 用于T2L的E13:G20（NETTO尺寸）
  headers.duzina = nettoDuzina;
  headers.sirina = nettoSirina;
  headers.visina = nettoVisina;
  
  console.log('文件1 - 完整headers对象:', {
    blkNo: headers.blkNo,
    cate: headers.cate,
    bruttoDuzina, bruttoSirina, bruttoVisina, bruttoM3,
    nettoDuzina, nettoSirina, nettoVisina, nettoM3,
    wgt, unitPrice: headers.unitPrice, totalPrice, kupac
  });
  headers.wgt = wgt;
  headers.totalPrice = totalPrice;
  
  // 解析数据行
  const stoneData = {};
  let dataRowCount = 0;
  
  console.log(`文件1 - 开始解析数据行，从第${headerRowIndex + 1}行到第${worksheet.rowCount}行`);
  
  // 从表头的下一行开始遍历数据
  for (let rowNum = headerRowIndex + 1; rowNum <= worksheet.rowCount; rowNum++) {
    const row = worksheet.getRow(rowNum);
    
    // 读取BLK NO的两列（数字部分 + 年份部分）
    const blkNoNumberCell = headers.blkNo ? row.getCell(headers.blkNo) : null;
    const blkNoYearCell = headers.blkNo ? row.getCell(headers.blkNo + 1) : null;
    
    const blkNoNumber = blkNoNumberCell ? blkNoNumberCell.value : null;
    const blkNoYear = blkNoYearCell ? blkNoYearCell.value : '';
    
    console.log(`文件1 - 第${rowNum}行, BLK NO第${headers.blkNo}列=${blkNoNumber}, 第${headers.blkNo + 1}列=${blkNoYear}`);
    
    if (!blkNoNumber) {
      console.log(`文件1 - 第${rowNum}行BLK NO数字部分为空，跳过`);
      continue; // 跳过空行
    }
    
    const blkNoNumberStr = String(blkNoNumber).trim();
    const blkNoYearStr = String(blkNoYear || '').trim();
    
    if (!blkNoNumberStr) {
      console.log(`文件1 - 第${rowNum}行BLK NO为空字符串，跳过`);
      continue;
    }
    
    // 跳过表头行（如果BLK NO包含"BLK"或"荒料号"等表头文字）
    if (blkNoNumberStr.toUpperCase().includes('BLK') || 
        blkNoNumberStr.includes('荒料号') || 
        blkNoNumberStr.toUpperCase().includes('BLOCK')) {
      console.log(`文件1 - 第${rowNum}行是表头，跳过`);
      continue;
    }
    
    dataRowCount++;
    
    // 拼接完整的BLK NO（如 "0205" + "/25" = "0205/25"）
    const fullBlkNo = blkNoNumberStr + blkNoYearStr;
    
    console.log(`文件1 - 第${rowNum}行, 完整BLK NO: ${fullBlkNo}`);
    
    // 使用完整的BLK NO作为key（如 "0205/25"）
    // 辅助函数：从单元格值中提取数字（处理公式单元格）
    const extractNumber = (cellValue) => {
      if (!cellValue) return 0;
      // 如果是对象（包含公式），取result属性
      if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        return parseFloat(cellValue.result) || 0;
      }
      // 否则直接转换
      return parseFloat(cellValue) || 0;
    };
    
    // 读取所有数值列
    const duzinaRaw = headers.duzina !== -1 ? row.getCell(headers.duzina).value : null;
    const sirinaRaw = headers.sirina !== -1 ? row.getCell(headers.sirina).value : null;
    const visinaRaw = headers.visina !== -1 ? row.getCell(headers.visina).value : null;
    const wgtRaw = headers.wgt !== -1 ? row.getCell(headers.wgt).value : null;
    const unitPriceRaw = headers.unitPrice ? row.getCell(headers.unitPrice).value : null;
    const totalPriceRaw = headers.totalPrice !== -1 ? row.getCell(headers.totalPrice).value : null;
    
    // 读取BRUTTO尺寸
    const bruttoDuzinaRaw = headers.bruttoDuzina !== -1 ? row.getCell(headers.bruttoDuzina).value : null;
    const bruttoSirinaRaw = headers.bruttoSirina !== -1 ? row.getCell(headers.bruttoSirina).value : null;
    const bruttoVisinaRaw = headers.bruttoVisina !== -1 ? row.getCell(headers.bruttoVisina).value : null;
    const bruttoM3Raw = headers.bruttoM3 !== -1 ? row.getCell(headers.bruttoM3).value : null;
    
    // 读取NETTO尺寸
    const nettoDuzinaRaw = headers.nettoDuzina !== -1 ? row.getCell(headers.nettoDuzina).value : null;
    const nettoSirinaRaw = headers.nettoSirina !== -1 ? row.getCell(headers.nettoSirina).value : null;
    const nettoVisinaRaw = headers.nettoVisina !== -1 ? row.getCell(headers.nettoVisina).value : null;
    const nettoM3Raw = headers.nettoM3 !== -1 ? row.getCell(headers.nettoM3).value : null;
    
    // 读取KUPAC（分两列，如 "HW" + "-" + "1" = "HW-1"）
    const kupacPart1Raw = headers.kupac !== -1 ? row.getCell(headers.kupac).value : '';
    const kupacPart2Raw = headers.kupac !== -1 ? row.getCell(headers.kupac + 1).value : '';
    
    const kupacPart1 = kupacPart1Raw ? String(kupacPart1Raw).trim() : '';
    const kupacPart2 = kupacPart2Raw ? String(kupacPart2Raw).trim() : '';
    
    // 拼接KUPAC（如果两部分都有值，用"-"连接）
    let kupacFull = '';
    if (kupacPart1 && kupacPart2) {
      kupacFull = `${kupacPart1}-${kupacPart2}`;
    } else if (kupacPart1) {
      kupacFull = kupacPart1;
    } else if (kupacPart2) {
      kupacFull = kupacPart2;
    }
    
    console.log(`文件1 - 第${rowNum}行 KUPAC: 列${headers.kupac}="${kupacPart1Raw}", 列${headers.kupac + 1}="${kupacPart2Raw}", 拼接结果="${kupacFull}"`);
    
    const totalPriceValue = extractNumber(totalPriceRaw);
    const unitPriceValue = extractNumber(unitPriceRaw);
    const nettoM3Value = extractNumber(nettoM3Raw);
    
    console.log(`文件1 - 第${rowNum}行 ${fullBlkNo}:`, {
      totalPrice: totalPriceValue,
      unitPrice: unitPriceValue,
      unitPriceRaw: unitPriceRaw,
      unitPriceCol: headers.unitPrice,
      nettoM3: nettoM3Value,
      nettoM3Raw: nettoM3Raw,
      nettoM3Col: headers.nettoM3,
      kupac: kupacFull
    });
    
    stoneData[fullBlkNo] = {
      blkNo: blkNoNumberStr,   // 数字部分：0205
      blkNoYear: blkNoYearStr, // 年份部分：/25
      fullBlkNo: fullBlkNo,    // 完整编号：0205/25
      cate: headers.cate ? (row.getCell(headers.cate).value || '') : '',
      // NETTO尺寸（兼容旧代码）
      duzina: extractNumber(duzinaRaw),
      sirina: extractNumber(sirinaRaw),
      visina: extractNumber(visinaRaw),
      // BRUTTO尺寸
      bruttoDuzina: extractNumber(bruttoDuzinaRaw),
      bruttoSirina: extractNumber(bruttoSirinaRaw),
      bruttoVisina: extractNumber(bruttoVisinaRaw),
      bruttoM3: extractNumber(bruttoM3Raw),
      // NETTO尺寸
      nettoDuzina: extractNumber(nettoDuzinaRaw),
      nettoSirina: extractNumber(nettoSirinaRaw),
      nettoVisina: extractNumber(nettoVisinaRaw),
      nettoM3: extractNumber(nettoM3Raw),
      // 其他
      wgt: extractNumber(wgtRaw),
      kupac: kupacFull,  // 拼接后的完整KUPAC（如 "HW 1"）
      unitPrice: unitPriceValue,
      totalPrice: totalPriceValue
    };
  }
  
  console.log(`文件1 - 共解析了${dataRowCount}行数据`);
  console.log('文件1解析结果（石头数据）:', stoneData);
  console.log('文件1 - stoneData的keys:', Object.keys(stoneData));
  
  return stoneData;
}

/**
 * 解析文件2 - 配柜方式
 * @param {File} file - Excel文件
 * @returns {Promise<Array>} - 柜数组 [{ctnNo, blockNrList}, ...]
 */
export async function parseContainerPlan(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0];
  
  let ctnColIndex = -1;
  let blockNrColIndex = -1;
  let headerRowIndex = -1;
  
  // 查找表头行，识别CTN列和Block nr.列
  console.log('文件2 - 开始查找表头...');
  
  // 遍历前10行，查找表头
  for (let rowNum = 1; rowNum <= Math.min(10, worksheet.rowCount); rowNum++) {
    if (ctnColIndex !== -1 && blockNrColIndex !== -1) break; // 已找到
    
    const row = worksheet.getRow(rowNum);
    console.log(`文件2 - 检查第${rowNum}行是否为表头`);
    
    // 遍历前15列
    for (let colNum = 1; colNum <= 15; colNum++) {
      const cell = row.getCell(colNum);
      const cellValue = cell.value || '';
      const cellStr = String(cellValue).toUpperCase().trim();
      
      if (colNum <= 12) { // 显示前12列的详细信息
        console.log(`文件2 - 第${rowNum}行第${colNum}列: "${cellValue}" (大写trim: "${cellStr}")`);
      }
      
      // 查找CTN列（方式1：直接查找包含CTN的列）
      if (ctnColIndex === -1 && cellStr && 
          (cellStr.includes('CTN') || cellStr.includes('CNTR'))) {
        // 如果包含CTN或CNTR，检查是否也包含NO，或者完全等于CTN/CNTR
        if (cellStr.includes('NO') || cellStr === 'CTN' || cellStr === 'CNTR') {
          ctnColIndex = colNum;
          headerRowIndex = rowNum;
          console.log(`文件2 - ✅ 找到CTN列（方式1）: 第${colNum}列，表头行: 第${rowNum}行`);
        }
      }
      
      // 查找Block nr.列
      if (blockNrColIndex === -1 && cellStr &&
          (cellStr.includes('BLK') || cellStr.includes('BLOCK') || cellStr.includes('BR.'))) {
        blockNrColIndex = colNum;
        headerRowIndex = rowNum;
        console.log(`文件2 - ✅ 找到Block nr.列: 第${colNum}列，表头行: 第${rowNum}行`);
        
        // 方式2：如果找到Block列，且CTN列还没找到，则Block列的前一列就是CTN列
        if (ctnColIndex === -1 && colNum > 1) {
          ctnColIndex = colNum - 1; // CTN列在Block列的前一列
          console.log(`文件2 - ✅ 推断CTN列（方式2）: 第${ctnColIndex}列（Block列的前一列）`);
        }
      }
    }
  }
  
  if (ctnColIndex === -1) {
    throw new Error('未找到CTN列（需要包含CTN NO./CTN/CNTR NO./CNTR）');
  }
  
  if (blockNrColIndex === -1) {
    throw new Error('未找到Block nr.列');
  }
  
  console.log(`开始解析配柜数据，CTN列索引: ${ctnColIndex}, Block nr.列索引: ${blockNrColIndex}`);
  
  // 按CTN值分组石头数据（处理合并单元格）
  const containerMap = new Map();
  let currentCtn = null;
  
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= headerRowIndex) return; // 跳过表头
    
    const ctnValue = ctnColIndex ? row.getCell(ctnColIndex).value : null;
    const blockNr = blockNrColIndex ? row.getCell(blockNrColIndex).value : null;
    
    // 如果CTN值非空，更新当前柜号
    if (ctnValue && String(ctnValue).trim()) {
      currentCtn = String(ctnValue).trim();
      if (!containerMap.has(currentCtn)) {
        containerMap.set(currentCtn, []);
      }
      console.log(`第${rowNumber}行: 发现新柜号 CTN=${currentCtn}`);
    }
    
    // 如果有Block nr.且有当前柜号，添加到对应的柜
    if (blockNr && String(blockNr).trim() && currentCtn) {
      const blockNrStr = String(blockNr).trim();
      containerMap.get(currentCtn).push(blockNrStr);
      console.log(`第${rowNumber}行: 添加石头 ${blockNrStr} 到柜${currentCtn}`);
    }
  });
  
  // 转换为数组格式
  const containers = [];
  containerMap.forEach((blockNrList, ctnNo) => {
    containers.push({
      ctnNo: ctnNo,
      blockNrList: blockNrList
    });
  });
  
  // 按CTN编号排序
  containers.sort((a, b) => {
    const numA = parseInt(a.ctnNo) || 0;
    const numB = parseInt(b.ctnNo) || 0;
    return numA - numB;
  });
  
  if (containers.length === 0) {
    throw new Error('未找到任何柜数据');
  }
  
  console.log(`文件2解析成功，找到 ${containers.length} 个柜:`, containers.map(c => `柜${c.ctnNo}(${c.blockNrList.length}颗石头)`).join(', '));
  
  // 详细显示每个柜的石头列表
  containers.forEach(container => {
    console.log(`柜${container.ctnNo}包含石头:`, container.blockNrList.join(', '));
  });
  
  return containers;
}

/**
 * 加载T2L模板
 * @param {File} file - Excel模板文件
 * @returns {Promise<Object>} - workbook对象
 */
export async function loadTemplate(file) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0];
  
  console.log('模板加载成功');
  console.log('模板工作表名称:', worksheet.name);
  console.log('模板行数:', worksheet.rowCount);
  console.log('模板列数:', worksheet.columnCount);
  
  return { workbook, worksheet };
}
