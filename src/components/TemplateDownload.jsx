import React from 'react';
import '../styles/TemplateDownload.css';

function TemplateDownload({ onBack }) {
  const templates = [
    {
      name: '文件1模板：block list with price',
      description: '包含BLK NO., CATE, 尺寸, 重量等石头所有信息',
      filename: 'block list with price.xlsx',
      icon: '📊'
    },
    {
      name: '文件2模板：配柜表(combination)',
      description: '配柜方式表，详细说明哪些石头装入哪个柜',
      filename: 'combination.xlsx',
      icon: '📦'
    },
    {
      name: '文件3模板：T2L',
      description: 'T2L模板，用于生成最终的交付单据',
      filename: 'T2L.xlsx',
      icon: '📄'
    },
    {
      name: '文件4模板：VGM',
      description: 'VGM模板，包含CNTR NO.、SEAL NO.等信息',
      filename: 'VGM.xlsx',
      icon: '🚢'
    }
  ];

  const handleDownload = (filename) => {
    // 创建下载链接
    const link = document.createElement('a');
    link.href = `/templates/${filename}`;
    link.download = filename;
    link.click();
  };

  return (
    <div className="template-download-container">
      <div className="template-header">
        <h2>📁 文件模板下载</h2>
        <p className="template-subtitle">下载以下Excel模板文件，填写后上传生成T2L</p>
      </div>

      <div className="template-list">
        {templates.map((template, index) => (
          <div key={index} className="template-card">
            <div className="template-icon">{template.icon}</div>
            <div className="template-info">
              <h3>{template.name}</h3>
              <p>{template.description}</p>
            </div>
            <button 
              className="download-btn"
              onClick={() => handleDownload(template.filename)}
            >
              ⬇️ 下载
            </button>
          </div>
        ))}
      </div>

      <div className="template-actions">
        <button className="back-btn" onClick={onBack}>
          ← 返回主界面
        </button>
      </div>

      <div className="template-tips">
        <h3>💡 使用提示</h3>
        <ul>
          <li>下载所有4个模板文件</li>
          <li>按照模板格式填写数据（不要修改表头）</li>
          <li>返回主界面上传文件并生成T2L</li>
        </ul>
      </div>
    </div>
  );
}

export default TemplateDownload;
