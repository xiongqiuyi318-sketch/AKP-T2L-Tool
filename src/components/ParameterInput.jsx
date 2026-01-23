import './ParameterInput.css';

export default function ParameterInput({ 
  startNumber, 
  setStartNumber, 
  year, 
  setYear,
  containerCount 
}) {
  const currentYear = new Date().getFullYear();

  return (
    <div className="parameter-input">
      <h3>T2L参数设置</h3>
      <div className="input-group">
        <div className="input-field">
          <label>T2L起始序号</label>
          <input 
            type="number" 
            min="1"
            value={startNumber}
            onChange={(e) => setStartNumber(parseInt(e.target.value) || 1)}
            placeholder="例如: 1"
          />
        </div>
        <div className="input-field">
          <label>年份</label>
          <input 
            type="number" 
            min="2000"
            max="2100"
            value={year}
            onChange={(e) => setYear(parseInt(e.target.value) || currentYear)}
            placeholder={`例如: ${currentYear}`}
          />
        </div>
      </div>
      {containerCount > 0 && (
        <div className="preview-info">
          <p>📦 检测到 <strong>{containerCount}</strong> 个集装箱</p>
          <p>📋 将生成T2L编号: <strong>{startNumber}</strong> 到 <strong>{startNumber + containerCount - 1}</strong></p>
        </div>
      )}
    </div>
  );
}
