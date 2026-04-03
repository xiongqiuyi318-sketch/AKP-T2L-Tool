import { useState } from 'react';
import FileUploader from './components/FileUploader';
import ParameterInput from './components/ParameterInput';
import TemplateDownload from './components/TemplateDownload';
import { parseStoneInfo, parseContainerPlan, loadTemplate } from './utils/excelParser';
import { matchStoneData } from './utils/dataProcessor';
import { buildWorkbookWithSheets, downloadExcel } from './utils/excelWriter';
import { generatePackingList, downloadPackingList } from './utils/packingListWriter';
import { parseVgmFile, generatePlWithCtnNoFromPacking, downloadPlWithCtnNo } from './utils/vgmWriter';
import { deleteCloudFile, fetchCustomerFiles, getSignedDownloadUrl, uploadBufferToCloud, uploadFileToCloud } from './utils/cloudApi';
import './App.css';

function App() {
  const [showTemplates, setShowTemplates] = useState(false);
  const [stoneInfo, setStoneInfo] = useState(null);
  const [containers, setContainers] = useState(null);
  const [template, setTemplate] = useState(null);
  const [vgmData, setVgmData] = useState(null);
  const [startNumber, setStartNumber] = useState(1);
  const [year, setYear] = useState(new Date().getFullYear());
  const [isGenerating, setIsGenerating] = useState(false);
  const [progress, setProgress] = useState('');
  const [customerName, setCustomerName] = useState('');
  const [batchCode, setBatchCode] = useState('');
  const [cloudEnabled, setCloudEnabled] = useState(true);
  const [cloudFiles, setCloudFiles] = useState([]);
  const [legacyFiles, setLegacyFiles] = useState([]);
  const [recentBatches, setRecentBatches] = useState([]);

  const normalizeText = (value) => String(value || '').trim();
  const toSlug = (value) =>
    String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/g, '-')
      .replace(/^-+|-+$/g, '') || 'default';
  const toSafeFileNameSegment = (value, fallback) =>
    String(value || '')
      .trim()
      .replace(/[\\/:*?"<>|]/g, '-')
      .replace(/\s+/g, '')
      .replace(/-+/g, '-')
      .replace(/^-+|-+$/g, '') || fallback;
  const getDisplayFileName = (pathname) => String(pathname || '').split('/').pop() || pathname;

  const validateBatchCode = () => {
    if (!cloudEnabled) return true;
    if (!batchCode.trim()) {
      alert('请填写批次，例如：第三批货');
      return false;
    }
    const normalized = normalizeText(batchCode);
    if (normalized !== batchCode) {
      setBatchCode(normalized);
    }
    return true;
  };

  const refreshCloudData = async (customer, batch = batchCode) => {
    if (!customer || !batch) return;
    try {
      const data = await fetchCustomerFiles(toSlug(customer), toSlug(batch));
      setCloudFiles(data.files);
      setLegacyFiles(data.legacyFiles);
      setRecentBatches(data.batches);
    } catch (error) {
      console.warn('刷新云端数据失败:', error);
    }
  };

  const saveInputIfEnabled = async (file, fileType) => {
    if (!cloudEnabled || !customerName.trim()) return;
    if (!validateBatchCode()) return;
    try {
      await uploadFileToCloud({
        file,
        customer: toSlug(customerName),
        batchCode: toSlug(batchCode),
        operator: 'unknown',
        fileType,
        kind: 'inputs'
      });
      await refreshCloudData(customerName, batchCode);
    } catch (error) {
      console.warn('云端归档失败（不影响本地解析）:', error);
    }
  };

  const saveOutputIfEnabled = async (workbook, fileType, filename) => {
    if (!cloudEnabled || !customerName.trim()) return;
    if (!validateBatchCode()) return;
    try {
      const buffer = await workbook.xlsx.writeBuffer();
      await uploadBufferToCloud({
        buffer,
        filename,
        customer: toSlug(customerName),
        batchCode: toSlug(batchCode),
        operator: 'unknown',
        fileType
      });
      await refreshCloudData(customerName, batchCode);
    } catch (error) {
      console.warn('云端保存失败（不影响本地下载）:', error);
      alert(`云端保存失败，但本地文件仍会下载。\n原因: ${error.message}`);
    }
  };

  const handleCloudFileDownload = async (pathname) => {
    try {
      const downloadUrl = await getSignedDownloadUrl(pathname);
      window.open(downloadUrl, '_blank', 'noopener,noreferrer');
    } catch (error) {
      alert(`获取下载链接失败: ${error.message}`);
    }
  };

  const handleCloudFileDelete = async (pathname) => {
    if (!customerName.trim()) {
      alert('请先填写客户名录，再执行删除');
      return;
    }
    const ok = confirm(`确定删除云端文件？\n${pathname}`);
    if (!ok) return;
    try {
      await deleteCloudFile({ pathname, customer: toSlug(customerName), operator: 'unknown' });
      await refreshCloudData(customerName, batchCode);
      alert('删除成功');
    } catch (error) {
      alert(`删除失败: ${error.message}`);
    }
  };

  const handleFile1Upload = async (file) => {
    const data = await parseStoneInfo(file);
    setStoneInfo(data);
    await saveInputIfEnabled(file, 'file1');
    console.log('文件1解析结果（石头数据）:', data);
  };

  const handleFile2Upload = async (file) => {
    const containerData = await parseContainerPlan(file);
    setContainers(containerData);
    await saveInputIfEnabled(file, 'file2');
  };

  const handleFile3Upload = async (file) => {
    const templateData = await loadTemplate(file);
    setTemplate(templateData);
    await saveInputIfEnabled(file, 'file3');
  };

  const handleVgmUpload = async (file) => {
    const parsedVgm = await parseVgmFile(file);
    setVgmData(parsedVgm);
    await saveInputIfEnabled(file, 'vgm');
    console.log('VGM解析结果:', parsedVgm);
  };

  const handleGenerate = async () => {
    if (!stoneInfo || !containers || !template) {
      alert('请先上传所有必需的文件！');
      return;
    }

    setIsGenerating(true);
    setProgress('正在处理数据...');

    try {
      // 为每个柜匹配石头数据
      const containersWithData = containers.map((container, index) => {
        setProgress(`正在处理第 ${index + 1}/${containers.length} 个柜...`);
        
        const { matchedStones, unmatchedBlocks } = matchStoneData(
          stoneInfo, 
          container.blockNrList
        );

        if (unmatchedBlocks.length > 0) {
          console.warn(`柜 ${container.ctnNo} 中有 ${unmatchedBlocks.length} 个石头未找到匹配:`, unmatchedBlocks);
        }

        if (matchedStones.length > 8) {
          console.warn(`柜 ${container.ctnNo} 包含 ${matchedStones.length} 颗石头，超过8颗限制，只会填充前8颗`);
        }

        // 调试：显示匹配的石头数据
        console.log(`柜 ${container.ctnNo} 匹配的石头:`, matchedStones);
        console.log(`柜 ${container.ctnNo} 总价计算:`, matchedStones.map(s => ({
          blkNo: s.blkNo,
          totalPrice: s.totalPrice,
          unitPrice: s.unitPrice,
          wgt: s.wgt
        })));

        return {
          ctnNo: container.ctnNo,
          matchedStones
        };
      });

      setProgress('正在生成Excel文件...');

      // 生成包含所有工作表的Excel文件
      const workbook = buildWorkbookWithSheets(
        template.workbook,
        template.worksheet,
        containersWithData,
        startNumber,
        year
      );

      setProgress('正在下载文件...');

      const outputFilename = `T2L-${toSafeFileNameSegment(customerName, 'customer')}-${toSafeFileNameSegment(batchCode, 'batch')}-${containers.length}柜.xlsx`;
      await saveOutputIfEnabled(workbook, 't2l', outputFilename);
      await downloadExcel(workbook, outputFilename);

      setProgress('');
      alert(`成功生成 ${containers.length} 个T2L工作表！`);
    } catch (error) {
      alert(`生成失败: ${error.message}`);
      console.error(error);
    } finally {
      setIsGenerating(false);
      setProgress('');
    }
  };

  const handleGeneratePackingList = async () => {
    if (!stoneInfo || !containers) {
      alert('请先上传文件1和文件2');
      return;
    }

    setIsGenerating(true);
    setProgress('正在生成Packing List...');

    try {
      // 为每个柜匹配石头数据
      const containersWithData = containers.map((container, index) => {
        const { matchedStones } = matchStoneData(
          stoneInfo, 
          container.blockNrList
        );

        return {
          ctnNo: container.ctnNo,
          matchedStones
        };
      });

      setProgress('正在生成Excel文件...');

      // 生成Packing List
      const workbook = await generatePackingList(containersWithData);

      setProgress('正在下载文件...');

      const outputFilename = `PL-${toSafeFileNameSegment(customerName, 'customer')}-${toSafeFileNameSegment(batchCode, 'batch')}-${containers.length}柜.xlsx`;
      await saveOutputIfEnabled(workbook, 'packing-list', outputFilename);
      await downloadPackingList(workbook, containers.length, outputFilename);

      setProgress('');
      alert(`Packing List 生成成功！包含 ${containers.length} 个柜`);
    } catch (error) {
      alert(`生成失败: ${error.message}`);
      console.error(error);
    } finally {
      setIsGenerating(false);
      setProgress('');
    }
  };

  const handleGeneratePlWithCtnNo = async () => {
    if (!stoneInfo || !containers || !vgmData) {
      alert('请先上传文件1、文件2和VGM文件');
      return;
    }

    setIsGenerating(true);
    setProgress('正在生成PL WITH CTN NO...');

    try {
      const containersWithData = containers.map((container) => {
        const { matchedStones } = matchStoneData(
          stoneInfo,
          container.blockNrList
        );

        return {
          ctnNo: container.ctnNo,
          matchedStones
        };
      });

      const workbook = await generatePlWithCtnNoFromPacking(containersWithData, vgmData);
      const outputFilename = `PL WITH CTN-${toSafeFileNameSegment(customerName, 'customer')}-${toSafeFileNameSegment(batchCode, 'batch')}-${containers.length}柜.xlsx`;
      await saveOutputIfEnabled(workbook, 'pl-with-ctn', outputFilename);
      await downloadPlWithCtnNo(workbook, outputFilename);
      setProgress('');
      alert('PL WITH CTN NO. 生成成功！');
    } catch (error) {
      alert(`生成失败: ${error.message}`);
      console.error(error);
    } finally {
      setIsGenerating(false);
      setProgress('');
    }
  };

  const isReadyToGenerate = stoneInfo && containers && template && startNumber > 0 && year > 0;
  const isReadyForPackingList = stoneInfo && containers;
  const isReadyForPlWithCtn = !!stoneInfo && !!containers && !!vgmData;

  // 如果显示模板页面，则渲染TemplateDownload组件
  if (showTemplates) {
    return (
      <div className="app">
        <header className="app-header">
          <h1>🚢 T2L自动生成工具</h1>
          <p>克罗地亚出口预发票自动化生成系统</p>
        </header>
        <main className="app-main">
          <TemplateDownload onBack={() => setShowTemplates(false)} />
        </main>
        <footer className="app-footer">
          <p>© 2026 AKP T2L自动化工具 | 专为克罗地亚石材出口设计</p>
        </footer>
      </div>
    );
  }

  // 主界面
  return (
    <div className="app">
      <header className="app-header">
        <h1>🚢 T2L自动生成工具</h1>
        <p>克罗地亚出口预发票自动化生成系统</p>
        <button 
          className="template-btn"
          onClick={() => setShowTemplates(true)}
        >
          📁 文件模板
        </button>
      </header>

      <main className="app-main">
        <section className="upload-section">
          <h2>📂 文件上传</h2>
          
          <FileUploader 
            label="文件1：block list with price (包含BLK NO., CATE, 尺寸, 重量等石头所有信息)"
            onFileSelect={handleFile1Upload}
          />

          <FileUploader 
            label="文件2: 配柜方式表 (包含Block nr.和CTN编号)"
            onFileSelect={handleFile2Upload}
          />

          <FileUploader 
            label="文件3: T2L模板 (空白T2L表格模板)"
            onFileSelect={handleFile3Upload}
          />

          <FileUploader
            label="文件4: VGM (用于生成PL WITH CTN NO.)"
            onFileSelect={handleVgmUpload}
          />
        </section>

        <section className="parameter-section">
          <ParameterInput 
            startNumber={startNumber}
            setStartNumber={setStartNumber}
            year={year}
            setYear={setYear}
            containerCount={containers?.length || 0}
          />
        </section>
        
        <section className="cloud-section">
          <h2>☁️ 云端归档（可选）</h2>
          <div className="cloud-controls">
            <label>
              客户名录
              <input
                type="text"
                value={customerName}
                onChange={(e) => setCustomerName(e.target.value)}
                placeholder="例如：HW / AKP"
              />
            </label>
            <label>
              批次编号
              <input
                type="text"
                value={batchCode}
                onChange={(e) => setBatchCode(normalizeText(e.target.value))}
                placeholder="例如：第三批货"
              />
            </label>
            <label className="checkbox-line">
              <input
                type="checkbox"
                checked={cloudEnabled}
                onChange={(e) => setCloudEnabled(e.target.checked)}
              />
              启用云端保存（上传与生成结果都会归档）
            </label>
            <button
              className="refresh-btn"
              onClick={() => {
                if (!validateBatchCode()) return;
                refreshCloudData(customerName, batchCode);
              }}
              disabled={!customerName.trim() || !batchCode.trim()}
            >
              刷新客户文件
            </button>
          </div>
          {recentBatches.length > 0 && (
            <div className="batch-switch">
              <span>最近批次：</span>
              <select
                value={batchCode}
                onChange={(e) => {
                  setBatchCode(e.target.value);
                  refreshCloudData(customerName, e.target.value);
                }}
              >
                <option value="">请选择批次</option>
                {recentBatches.map((batch) => (
                  <option key={batch} value={batch}>{batch}</option>
                ))}
              </select>
            </div>
          )}
          <div className="cloud-lists">
            <div>
              <h3>客户文件（{customerName || '-'} / {batchCode || '-'}）</h3>
              <ul>
                {cloudFiles.slice(0, 8).map((file) => (
                  <li key={file.pathname}>
                    <div className="file-item">
                      <button
                        type="button"
                        className="file-link-btn"
                        onClick={() => handleCloudFileDownload(file.pathname)}
                        title={file.pathname}
                      >
                        {getDisplayFileName(file.pathname)}
                      </button>
                      <button
                        type="button"
                        className="file-del-btn"
                        onClick={() => handleCloudFileDelete(file.pathname)}
                        title="删除"
                      >
                        删除
                      </button>
                    </div>
                  </li>
                ))}
              </ul>
              {legacyFiles.length > 0 && (
                <>
                  <h3>旧结构文件（兼容）</h3>
                  <ul>
                    {legacyFiles.slice(0, 8).map((file) => (
                      <li key={`legacy-${file.pathname}`}>
                        <div className="file-item">
                          <button
                            type="button"
                            className="file-link-btn"
                            onClick={() => handleCloudFileDownload(file.pathname)}
                            title={file.pathname}
                          >
                            {getDisplayFileName(file.pathname)}
                          </button>
                        </div>
                      </li>
                    ))}
                  </ul>
                </>
              )}
            </div>
          </div>
        </section>

        <section className="action-section">
          <button 
            className="generate-btn"
            onClick={handleGenerate}
            disabled={!isReadyToGenerate || isGenerating}
          >
            {isGenerating ? '生成中...' : '🚀 生成T2L文件'}
          </button>
          
          <button 
            className="packing-list-btn"
            onClick={handleGeneratePackingList}
            disabled={!isReadyForPackingList || isGenerating}
          >
            {isGenerating ? '生成中...' : '📦 生成 Packing List'}
          </button>

          <button
            className="pl-ctn-btn"
            onClick={handleGeneratePlWithCtnNo}
            disabled={!isReadyForPlWithCtn || isGenerating}
          >
            {isGenerating ? '生成中...' : '🧾 生成 PL WITH CTN NO.'}
          </button>
          
          {progress && (
            <div className="progress-info">
              {progress}
            </div>
          )}
        </section>

        <section className="info-section">
          <h3>使用说明</h3>
          <ol>
            <li>上传<strong>文件1</strong>（石头信息表），包含所有石头的详细信息</li>
            <li>上传<strong>文件2</strong>（配柜方式表），说明每个柜包含哪些石头</li>
            <li>上传<strong>文件3</strong>（T2L模板），作为生成T2L的基础模板</li>
            <li>设置<strong>T2L起始序号</strong>和<strong>年份</strong></li>
            <li>点击"生成T2L文件"按钮，系统将自动生成一个包含多个工作表的Excel文件</li>
            <li>每个工作表对应一个集装箱，工作表名为CTN编号</li>
            <li>上传<strong>文件1、文件2、VGM文件</strong>后，可基于Packing List明细生成"PL WITH CTN NO."</li>
          </ol>
        </section>
      </main>

      <footer className="app-footer">
        <p>© 2026 AKP T2L自动化工具 | 专为克罗地亚石材出口设计</p>
      </footer>
    </div>
  );
}

export default App;
