/**
 * 根据Block nr.列表匹配石头数据
 * @param {Object} stoneInfo - 石头数据字典
 * @param {Array<string>} blockNrList - Block nr.列表
 * @returns {Object} - {matchedStones: Array, unmatchedBlocks: Array}
 */
export function matchStoneData(stoneInfo, blockNrList) {
  const matchedStones = [];
  const unmatchedBlocks = [];
  
  // 调试：显示stoneInfo的所有key
  console.log('stoneInfo的所有key:', Object.keys(stoneInfo));
  console.log('要匹配的blockNrList:', blockNrList);
  
  blockNrList.forEach(blockNr => {
    // 尝试多种匹配方式
    const blockNrStr = String(blockNr).trim();
    
    // 方式1：直接匹配
    let stone = stoneInfo[blockNrStr];
    
    // 方式2：如果blockNr包含年份后缀（如0208/25），尝试只用数字部分匹配
    if (!stone && blockNrStr.includes('/')) {
      const blockNrOnly = blockNrStr.split('/')[0].trim();
      stone = stoneInfo[blockNrOnly];
      console.log(`尝试用数字部分 ${blockNrOnly} 匹配 ${blockNrStr}:`, stone ? '成功' : '失败');
    }
    
    // 方式3：如果blockNr不包含年份，尝试在stoneInfo中查找带年份的key
    if (!stone) {
      const matchingKey = Object.keys(stoneInfo).find(key => key.startsWith(blockNrStr));
      if (matchingKey) {
        stone = stoneInfo[matchingKey];
        console.log(`尝试前缀匹配 ${blockNrStr} -> ${matchingKey}:`, stone ? '成功' : '失败');
      }
    }
    
    if (stone) {
      matchedStones.push(stone);
    } else {
      unmatchedBlocks.push(blockNrStr);
      console.log(`未找到匹配: ${blockNrStr}`);
    }
  });
  
  return {
    matchedStones,
    unmatchedBlocks
  };
}

/**
 * 计算总价值
 * @param {Array} matchedStones - 匹配的石头数据数组
 * @returns {number} - 总价值
 */
export function calculateTotalValue(matchedStones) {
  return matchedStones.reduce((sum, stone) => {
    // 优先使用totalPrice字段，如果没有则计算unitPrice × wgt
    const totalPrice = parseFloat(stone.totalPrice) || 0;
    if (totalPrice > 0) {
      return sum + totalPrice;
    } else {
      const unitPrice = parseFloat(stone.unitPrice) || 0;
      const weight = parseFloat(stone.wgt) || 0;
      return sum + (unitPrice * weight);
    }
  }, 0);
}

/**
 * 格式化B10单元格文本
 * @param {number} t2lNumber - T2L编号
 * @param {number} year - 年份
 * @param {number} containerIndex - 柜序号（从1开始）
 * @returns {string} - 格式化的多行文本
 */
export function formatB10Text(t2lNumber, year, containerIndex) {
  const yearSuffix = String(year).slice(-2); // 取年份后两位
  return `T2L\nDELIVERY NOTE - ${t2lNumber} /${yearSuffix}\nDATE(DATUM):\n${containerIndex} KAMION`;
}
