import { useState } from 'react';
import FileUploader from './components/FileUploader';
import ParameterInput from './components/ParameterInput';
import TemplateDownload from './components/TemplateDownload';
import { parseStoneInfo, parseContainerPlan, loadTemplate } from './utils/excelParser';
import { matchStoneData } from './utils/dataProcessor';
import { buildWorkbookWithSheets, downloadExcel } from './utils/excelWriter';
import './App.css';

function App() {
  const [showTemplates, setShowTemplates] = useState(false);
  const [stoneInfo, setStoneInfo] = useState(null);
  const [containers, setContainers] = useState(null);
  const [template, setTemplate] = useState(null);
  const [startNumber, setStartNumber] = useState(1);
  const [year, setYear] = useState(new Date().getFullYear());
  const [isGenerating, setIsGenerating] = useState(false);
  const [progress, setProgress] = useState('');

  const handleFile1Upload = async (file) => {
    const data = await parseStoneInfo(file);
    setStoneInfo(data);
    console.log('文件1解析结果（石头数据）:', data);
  };

  const handleFile2Upload = async (file) => {
    const containerData = await parseContainerPlan(file);
    setContainers(containerData);
  };

  const handleFile3Upload = async (file) => {
    const templateData = await loadTemplate(file);
    setTemplate(templateData);
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

      // 下载文件
      await downloadExcel(workbook, `T2L_Output_${startNumber}-${startNumber + containers.length - 1}.xlsx`);

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

  const isReadyToGenerate = stoneInfo && containers && template && startNumber > 0 && year > 0;

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

        <section className="action-section">
          <button 
            className="generate-btn"
            onClick={handleGenerate}
            disabled={!isReadyToGenerate || isGenerating}
          >
            {isGenerating ? '生成中...' : '🚀 生成T2L文件'}
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
