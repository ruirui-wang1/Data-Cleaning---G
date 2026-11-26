/**
 * ============================================================================
 * 通用跨表优先级匹配模板（Google Apps Script）
 * ============================================================================
 * 功能概述：
 *   1. 支持通过目标表“复选框（安装型 On edit 触发器）”或“顶部菜单”触发匹配；
 *   2. 按 sources 数组顺序逐表查找匹配主键，先匹配先锁定（高优先级覆盖低优先级）；
 *   3. 命中全部目标主键后提前终止后续扫描，节省执行时间与配额；
 *   4. 按 chunkSize 分块扫描主键列，仅在批次命中时读取宽列数据；
 *   5. 内置 LockService 并发锁、状态单元格提示、Toast 提示与复选框自动复位。
 *
 * 使用步骤：
 *   1. 在目标工作表的 triggerRow 行、triggerCol 列（默认 J1）插入“复选框”；
 *   2. 修改下方 TEMPLATE_CONFIG 与 mapMatchedRow() 以适配你的业务字段；
 *   3. 将本脚本粘贴至“扩展程序 -> Apps Script”并保存；
 *   4. 在 Apps Script 左侧“触发器”中添加触发器：
 *        - 选择函数：onCheckboxTrigger
 *        - 事件来源：自电子表格
 *        - 事件类型：编辑时
 *   5. 首次运行需完成 Google 账号权限授权。
 * ============================================================================
 */

// ================= 一、 配置区（按业务需求修改） =================
const TEMPLATE_CONFIG = {
  // 目标工作表与读写位置配置
  target: {
    sheetName: 'JIT核价商家版',      // 目标工作表名称
    triggerRow: 1,                  // 复选框所在行号（从 1 开始）
    triggerCol: 10,                 // 复选框所在列号（10 = J列）
    statusCol: 11,                  // 运行状态显示列号（11 = K列，建议避开表头或单独占用）
    keyCol: 3,                      // 目标表匹配主键所在列号（3 = C列，从第 2 行开始读取）
    outputStartCol: 5,              // 结果写入起始列号（5 = E列，从第 2 行开始写入）
    outputNumCols: 5,               // 结果写入列数（5 表示 E:I 共 5 列；须与 mapMatchedRow 返回长度一致）
    menuTitle: '⚡ 跨表数据匹配',     // 顶部菜单名称
    menuItemLabel: '开始匹配数据'   // 顶部菜单项名称
  },

  // 数据源列表（严格按匹配优先级从高到低排列；前面匹配到的主键不会被后面覆盖）
  sources: [
    {
      label: '12月',
      id: '1rSwtmXQXIO45gmUtVMQcbKT2aYaYntoPrYpSAiz8NXM',
      sheetName: 'Seller nomi JIT提报表-LS',
      gid: 1313946005
    },
    {
      label: '11月',
      id: '1cGP9TgQhYl6N7b1P2qjuhAAYvWVkgIMa2DF5KLtLrac',
      sheetName: 'Seller nomi JIT提报表-LS',
      gid: 1313946005
    },
    {
      label: '10月',
      id: '1lwKSNlq4jMlWkYamQIp5qwFyJbmm0uIUeXMpF3RZ2oM',
      sheetName: 'Seller nomi JIT提报表-LS',
      gid: 1313946005
    }
  ],

  // 源表读取范围与性能调优参数
  sourceRead: {
    keyCol: 4,                      // 源表主键所在列号（4 = D列，从第 2 行开始扫描）
    blockStartCol: 20,              // 源表宽数据块起始列号（20 = T列）
    blockNumCols: 12,               // 源表宽数据块列数（12 表示 T:AE 共 12 列）
    chunkSize: 3000,                // 单次分块读取行数（数据量大或易超时可适当调小，如 1500~3000）
    lockWaitMs: 1000                // 并发锁等待毫秒数
  }
};

/**
 * 动态字段映射规则（核心业务定制区）
 * @param {Array<*>} rowData - 从 blockStartCol 起算、长度为 blockNumCols 的一维行数组。
 *                             列下标换算公式：index = 实际列号 - TEMPLATE_CONFIG.sourceRead.blockStartCol
 *                             例如 blockStartCol = 20（T列）：
 *                               - index 0  -> 第 20 列（T列）
 *                               - index 10 -> 第 30 列（AD列）
 *                               - index 11 -> 第 31 列（AE列）
 * @param {string} sourceLabel - 当前命中的数据源标签（如 '12月'）
 * @returns {Array<*>} 返回长度等于 TEMPLATE_CONFIG.target.outputNumCols 的一维输出行
 */
function mapMatchedRow(rowData, sourceLabel) {
  const valAD = rowData[10] !== undefined ? rowData[10] : '';
  const valAE = rowData[11] !== undefined ? rowData[11] : '';
  const reason = String(valAE).trim();

  let colG = '';
  let colH = '';

  if (reason === '1688三档价不通过') {
    colG = rowData[3] !== undefined ? rowData[3] : '';
    colH = rowData[5] !== undefined ? rowData[5] : '';
  } else if (reason === '架上同款cogs不通过') {
    colG = rowData[8] !== undefined ? rowData[8] : '';
    colH = rowData[9] !== undefined ? rowData[9] : '';
  } else if (reason === '倒算价不通过') {
    colG = rowData[0] !== undefined ? rowData[0] : '';
    colH = rowData[1] !== undefined ? rowData[1] : '';
  }

  return [valAD, valAE, colG, colH, sourceLabel];
}

// ================= 二、 触发器与菜单入口 =================

/**
 * 打开电子表格时挂载自定义菜单（支持手动触发）
 */
function onOpen() {
  const cfg = TEMPLATE_CONFIG.target;
  SpreadsheetApp.getUi()
    .createMenu(cfg.menuTitle)
    .addItem(cfg.menuItemLabel, 'runMatchFromMenu')
    .addToUi();
}

/**
 * 安装型可编辑触发器入口（须在 Apps Script 触发器面板中手动绑定此函数）
 * 注意：函数名避免使用保留字 onEdit，防止触发权限受限的简单触发器
 * @param {Object} e - Google Sheets 编辑事件对象
 */
function onCheckboxTrigger(e) {
  if (!e || !e.range) return;

  const targetCfg = TEMPLATE_CONFIG.target;
  const sheet = e.range.getSheet();

  if (
    sheet.getName() !== targetCfg.sheetName ||
    e.range.getRow() !== targetCfg.triggerRow ||
    e.range.getColumn() !== targetCfg.triggerCol
  ) {
    return;
  }

  // 仅当复选框变为勾选状态（TRUE）时执行
  if (e.value !== 'TRUE' && e.range.getValue() !== true) {
    return;
  }

  executeCrossSheetMatch(true);
}

/**
 * 顶部菜单手动运行入口
 */
function runMatchFromMenu() {
  executeCrossSheetMatch(false);
}

// ================= 三、 核心匹配引擎（通常无需修改） =================

/**
 * 执行跨表优先级匹配主流程
 * @param {boolean} fromCheckbox - 是否由复选框触发
 */
function executeCrossSheetMatch(fromCheckbox) {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetCfg = TEMPLATE_CONFIG.target;
  const readCfg = TEMPLATE_CONFIG.sourceRead;
  const sources = TEMPLATE_CONFIG.sources;

  const priorityRangeText = buildPriorityRangeText(sources);
  const targetSheet = ss.getSheetByName(targetCfg.sheetName);

  if (!targetSheet) {
    const errMsg = '未找到目标工作表：' + targetCfg.sheetName;
    if (!fromCheckbox) {
      SpreadsheetApp.getUi().alert(errMsg);
    } else {
      ss.toast(errMsg, '❌ 配置错误', 8);
    }
    return;
  }

  const triggerCell = targetSheet.getRange(targetCfg.triggerRow, targetCfg.triggerCol);
  const statusCell = targetSheet.getRange(targetCfg.triggerRow, targetCfg.statusCol);

  // 1. 获取脚本锁，防止多人并发编辑导致写入冲突
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(readCfg.lockWaitMs)) {
    ss.toast('后台已有匹配任务正在执行，请勿重复勾选。', '⏳ 任务运行中', 5);
    if (fromCheckbox) {
      triggerCell.setValue(false);
    }
    return;
  }

  try {
    // 立即更新界面状态与提示
    statusCell.setValue('⏳ 正在匹配 (' + priorityRangeText + ')，请稍候...');
    ss.toast('正在按 ' + priorityRangeText + ' 跨表匹配数据，请稍候...', '⚡ 开始匹配', 8);
    SpreadsheetApp.flush();

    if (!Array.isArray(sources) || sources.length === 0) {
      throw new Error('数据源列表 sources 为空，请检查配置。');
    }

    // 2. 读取目标表主键列（第 2 行至最后一行）
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow < 2) {
      const msg = '⚠️ 目标表第 2 行以下没有需要匹配的数据';
      statusCell.setValue(msg);
      if (!fromCheckbox) SpreadsheetApp.getUi().alert(msg);
      return;
    }

    const targetKeysRaw = targetSheet.getRange(2, targetCfg.keyCol, targetLastRow - 1, 1).getValues();
    const neededKeys = new Set();
    const targetKeys = targetKeysRaw.map(function (row) {
      const key = normalizeKey(row[0]);
      if (key !== '') {
        neededKeys.add(key);
      }
      return key;
    });

    // 剔除尾部全空行，缩小写入范围
    let effectiveTargetRows = targetKeys.length;
    while (effectiveTargetRows > 0 && targetKeys[effectiveTargetRows - 1] === '') {
      effectiveTargetRows--;
    }

    if (neededKeys.size === 0 || effectiveTargetRows === 0) {
      const msg = '⚠️ 目标主键列内容均为空';
      statusCell.setValue(msg);
      if (!fromCheckbox) SpreadsheetApp.getUi().alert(msg);
      return;
    }

    // 3. 按优先级分块扫描各数据源
    const matchedMap = new Map();
    const chunkSize = Math.max(1, Number(readCfg.chunkSize) || 3000);

    for (let i = 0; i < sources.length; i++) {
      if (matchedMap.size >= neededKeys.size) break;

      const src = sources[i];
      const srcSheet = resolveSourceSheet(src);
      const srcLastRow = srcSheet.getLastRow();
      if (srcLastRow < 2) continue;

      for (let startRow = 2; startRow <= srcLastRow; startRow += chunkSize) {
        if (matchedMap.size >= neededKeys.size) break;

        const batchRows = Math.min(chunkSize, srcLastRow - startRow + 1);
        const keyValues = srcSheet.getRange(startRow, readCfg.keyCol, batchRows, 1).getValues();

        // 先扫描当前分块主键列，若无任何“所需且未匹配”的主键则跳过宽列读取
        let hasHitInBatch = false;
        for (let r = 0; r < batchRows; r++) {
          const key = normalizeKey(keyValues[r][0]);
          if (key !== '' && neededKeys.has(key) && !matchedMap.has(key)) {
            hasHitInBatch = true;
            break;
          }
        }

        if (!hasHitInBatch) continue;

        const blockValues = srcSheet
          .getRange(startRow, readCfg.blockStartCol, batchRows, readCfg.blockNumCols)
          .getValues();

        for (let r = 0; r < batchRows; r++) {
          const key = normalizeKey(keyValues[r][0]);
          if (key !== '' && neededKeys.has(key) && !matchedMap.has(key)) {
            const mappedRow = mapMatchedRow(blockValues[r], src.label);
            matchedMap.set(key, formatOutputRow(mappedRow, targetCfg.outputNumCols));

            if (matchedMap.size >= neededKeys.size) break;
          }
        }
      }
    }

    // 4. 组装结果矩阵并批量回写目标表输出列
    const emptyRow = createEmptyRow(targetCfg.outputNumCols);
    let matchedRowCount = 0;
    const outputData = [];

    for (let i = 0; i < effectiveTargetRows; i++) {
      const key = targetKeys[i];
      if (key !== '' && matchedMap.has(key)) {
        matchedRowCount++;
        outputData.push(matchedMap.get(key));
      } else {
        outputData.push(emptyRow.slice());
      }
    }

    targetSheet
      .getRange(2, targetCfg.outputStartCol, effectiveTargetRows, targetCfg.outputNumCols)
      .setValues(outputData);

    // 5. 输出执行摘要
    const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
    const timeNow = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
    const summaryText =
      '✅ 匹配完成：' +
      matchedRowCount +
      '/' +
      effectiveTargetRows +
      ' 行 (' +
      elapsedSec +
      '秒 | ' +
      timeNow +
      ')';

    statusCell.setValue(summaryText);
    ss.toast('成功匹配 ' + matchedRowCount + ' 行，耗时 ' + elapsedSec + ' 秒。', '✅ 匹配完成', 8);

    if (!fromCheckbox) {
      SpreadsheetApp.getUi().alert(
        '✅ 数据匹配完成！\n\n' +
          '• 匹配顺序：' + priorityRangeText + '\n' +
          '• 有效查询行数：' + effectiveTargetRows + ' 行\n' +
          '• 成功匹配行数：' + matchedRowCount + ' 行\n' +
          '• 执行耗时：' + elapsedSec + ' 秒'
      );
    }
  } catch (err) {
    statusCell.setValue('❌ 匹配出错：' + err.message);
    ss.toast(err.message, '❌ 匹配失败', 10);
    if (!fromCheckbox) {
      SpreadsheetApp.getUi().alert('❌ 运行出错：\n' + err.message);
    }
  } finally {
    // 无论成功或失败，均复位复选框并释放并发锁
    if (fromCheckbox) {
      try {
        triggerCell.setValue(false);
      } catch (resetErr) {
        // 忽略复位异常，避免掩盖主错误
      }
    }
    try {
      lock.releaseLock();
    } catch (lockErr) {
      // 忽略锁释放异常
    }
  }
}

// ================= 四、 辅助工具函数 =================

/**
 * 标准化主键文本（去首尾空格、统一转字符串）
 * @param {*} rawValue
 * @returns {string}
 */
function normalizeKey(rawValue) {
  if (rawValue === null || rawValue === undefined) return '';
  return String(rawValue).trim();
}

/**
 * 解析并打开数据源工作表（优先按 sheetName，其次按 gid）
 * @param {Object} src - 数据源配置项
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function resolveSourceSheet(src) {
  let srcSpreadsheet;
  try {
    srcSpreadsheet = SpreadsheetApp.openById(src.id);
  } catch (err) {
    throw new Error('无法打开【' + src.label + '】表格，请检查访问权限与表格 ID：' + err.message);
  }

  let srcSheet = null;
  if (src.sheetName) {
    srcSheet = srcSpreadsheet.getSheetByName(src.sheetName);
  }
  if (!srcSheet && src.gid !== undefined && src.gid !== null) {
    srcSheet = srcSpreadsheet.getSheets().find(function (s) {
      return s.getSheetId() === Number(src.gid);
    }) || null;
  }

  if (!srcSheet) {
    throw new Error(
      '在【' + src.label + '】表格中未找到目标工作表（名称：' + (src.sheetName || '未指定') + '）'
    );
  }

  return srcSheet;
}

/**
 * 生成优先级范围提示文案（如“12月→10月”）
 * @param {Array<Object>} sources
 * @returns {string}
 */
function buildPriorityRangeText(sources) {
  if (!Array.isArray(sources) || sources.length === 0) return '未配置数据源';
  if (sources.length === 1) return sources[0].label;
  return sources[0].label + '→' + sources[sources.length - 1].label;
}

/**
 * 确保单行输出数组长度严格等于 outputNumCols
 * @param {Array<*>} row
 * @param {number} numCols
 * @returns {Array<*>}
 */
function formatOutputRow(row, numCols) {
  const result = [];
  for (let i = 0; i < numCols; i++) {
    result.push(row && row[i] !== undefined && row[i] !== null ? row[i] : '');
  }
  return result;
}

/**
 * 创建指定列数的空行数组
 * @param {number} numCols
 * @returns {Array<string>}
 */
function createEmptyRow(numCols) {
  return new Array(numCols).fill('');
}

