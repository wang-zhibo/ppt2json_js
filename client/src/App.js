import React, { useState } from 'react';
import './App.css';
import Notification from './Notification';

function App() {
  const [file, setFile] = useState(null);
  const [result, setResult] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [notification, setNotification] = useState({ visible: false, message: '', type: 'info' });

  const showNotification = (message, type = 'info') => {
    setNotification({ visible: true, message, type });
    // 自动隐藏通知（除了错误类型）
    if (type !== 'error') {
      setTimeout(() => {
        setNotification(prev => ({ ...prev, visible: false }));
      }, 3000);
    }
  };

  const hideNotification = () => {
    setNotification(prev => ({ ...prev, visible: false }));
  };

  const handleFileChange = (e) => {
    console.log('handleFileChange 被调用');
    const selectedFile = e.target.files[0];
    console.log('选择的文件:', selectedFile);
    
    if (!selectedFile) {
      console.log('没有选择文件');
      return;
    }
    
    // 检查文件类型
    const isPPT = selectedFile.type === 'application/vnd.ms-powerpoint' || 
                  selectedFile.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
                  selectedFile.name.endsWith('.ppt') || 
                  selectedFile.name.endsWith('.pptx');
    
    console.log('文件类型检查:', {
      type: selectedFile.type,
      name: selectedFile.name,
      isPPT: isPPT
    });
    
    if (!isPPT) {
      console.log('文件类型不支持');
      showNotification('只支持PPT和PPTX文件格式！', 'warning');
      return;
    }
    
    // 检查文件大小 (50MB)
    const isLt50M = selectedFile.size / 1024 / 1024 < 50;
    console.log('文件大小检查:', {
      size: selectedFile.size,
      sizeMB: selectedFile.size / 1024 / 1024,
      isLt50M: isLt50M
    });
    
    if (!isLt50M) {
      console.log('文件大小超限');
      showNotification('文件大小不能超过50MB！', 'warning');
      return;
    }
    
    console.log('文件验证通过，设置文件状态');
    setFile(selectedFile);
    setResult(null);
    setError(null);
    showNotification(`${selectedFile.name} 文件已选择`, 'success');
  };

  const handleUpload = async () => {
    console.log('handleUpload 开始执行');
    console.log('当前文件:', file);
    
    if (!file) {
      console.log('没有选择文件');
      setError('请选择文件');
      return;
    }

    console.log('开始上传流程');
    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);
    console.log('FormData已创建，文件大小:', file.size);

    try {
      console.log('准备发送fetch请求到 /api/convert');
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      console.log('收到响应，状态码:', response.status);
      console.log('响应头:', Object.fromEntries(response.headers.entries()));

      const data = await response.json();
      console.log('响应数据:', data);

      if (response.ok) {
        console.log('请求成功，设置结果');
        setResult(data.data);
        console.log('解析结果:', data.data);
        showNotification('PPT文件转换成功！', 'success');
      } else {
        console.log('请求失败，设置错误');
        setError(data.error || '转换失败');
        showNotification(data.error || '转换失败', 'error');
      }
    } catch (err) {
      console.error('上传过程中发生错误:', err);
      console.error('错误详情:', {
        name: err.name,
        message: err.message,
        stack: err.stack
      });
      setError('网络错误: ' + err.message);
      showNotification('网络错误: ' + err.message, 'error');
    } finally {
      console.log('上传流程结束，设置loading为false');
      setLoading(false);
    }
  };

  const downloadJSON = () => {
    if (!result) return;

    const dataStr = JSON.stringify(result, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'ppt-data.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showNotification('JSON文件下载成功！', 'success');
  };

  const copyJSON = async () => {
    if (!result) return;

    const dataStr = JSON.stringify(result, null, 2);
    
    try {
      // 检查是否支持现代剪贴板API
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(dataStr);
        showNotification('JSON已复制到剪贴板！', 'success');
      } else {
        // 备用方案：使用传统的document.execCommand
        const textArea = document.createElement('textarea');
        textArea.value = dataStr;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        
        const successful = document.execCommand('copy');
        document.body.removeChild(textArea);
        
        if (successful) {
          showNotification('JSON已复制到剪贴板！', 'success');
        } else {
          throw new Error('复制命令执行失败');
        }
      }
    } catch (err) {
      console.error('复制失败:', err);
      // 最后的备用方案：显示JSON内容供用户手动复制
      alert('自动复制失败，请手动复制以下内容：\n\n' + dataStr);
      showNotification('复制失败，已显示内容供手动复制', 'error');
    }
  };

  return (
    <div className="App">
      <Notification
        visible={notification.visible}
        message={notification.message}
        type={notification.type}
        onClose={hideNotification}
      />
      
      <header className="App-header">
        <h1>PPT转JSON工具</h1>
        <p>上传PPT或PPTX文件，获取JSON格式的演示文稿数据</p>
      </header>

      <main className="App-main">
        <div className="upload-section">
          <h3>上传PPT文件</h3>
          <p>支持PPT和PPTX格式，最大文件大小50MB。转换后的JSON将包含幻灯片的所有元素信息。</p>
          
          <input
            type="file"
            accept=".ppt,.pptx"
            onChange={handleFileChange}
            className="file-input"
          />
          
          {file && (
            <div style={{ marginTop: '20px', textAlign: 'center' }}>
              <p style={{ color: '#e0e6ef', marginBottom: '10px' }}>
                已选择文件: <strong>{file.name}</strong>
              </p>
              <button 
                onClick={handleUpload}
                disabled={loading}
                className="upload-button"
              >
                {loading ? '转换中...' : '开始转换'}
              </button>
            </div>
          )}

          {loading && (
            <div className="loading-section">
              <p>正在转换PPT文件，请稍候...</p>
            </div>
          )}

          {error && (
            <div className="error-message">
              <h3>错误信息</h3>
              <p>{error}</p>
            </div>
          )}
        </div>

        {result && (
          <div className="result-section">
            <h3>转换结果</h3>
            
            <div className="result-actions">
              <button 
                onClick={copyJSON}
                className="action-button copy-button"
                title="复制JSON到剪贴板"
              >
                📋 复制JSON
              </button>
              <button 
                onClick={downloadJSON}
                className="action-button download-button"
                title="下载JSON文件"
              >
                💾 下载JSON文件
              </button>
            </div>
            
            <div className="result-summary">
              <h4>基本信息</h4>
              <p><strong>幻灯片数量:</strong> {result.slides.length}</p>
              <p><strong>主题颜色:</strong> {Object.keys(result.themeColors || {}).length} 种</p>
              <p><strong>尺寸:</strong> {result.width} x {result.height} (EMU)</p>
            </div>

            {result.slides.length > 0 ? (
              <div className="slides-info">
                <h4>幻灯片详情</h4>
                {result.slides.map((slide, index) => (
                  <div key={index} className="slide-info">
                    <h5>幻灯片 {index + 1}</h5>
                    <p><strong>元素数量:</strong> {slide.elements.length}</p>
                    {slide.error && <p className="slide-error">错误: {slide.error}</p>}
                    {slide.elements.length > 0 && (
                      <div className="elements-preview">
                        <h6>元素预览:</h6>
                        {slide.elements.slice(0, 3).map((element, elemIndex) => (
                          <div key={elemIndex} className="element-preview">
                            <span><strong>类型:</strong> {element.type}</span>
                            {element.content && <span><strong>内容:</strong> {element.content.substring(0, 50)}...</span>}
                          </div>
                        ))}
                        {slide.elements.length > 3 && <p>...还有 {slide.elements.length - 3} 个元素</p>}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ) : (
              <div className="no-slides">
                <h4>⚠️ 未找到幻灯片内容</h4>
                <p>这可能是因为:</p>
                <ul>
                  <li>文件格式不兼容</li>
                  <li>文件损坏</li>
                  <li>解析器需要更新</li>
                </ul>
              </div>
            )}

            <div className="json-output">
              <h4>完整JSON数据</h4>
              <details>
                <summary>点击查看完整JSON</summary>
                <pre>{JSON.stringify(result, null, 2)}</pre>
              </details>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

export default App; 
