import React, { useState } from 'react';
import { Layout, Typography, Card, Upload, Button, message, Spin, Alert } from 'antd';
import { UploadOutlined, FileTextOutlined, DownloadOutlined } from '@ant-design/icons';
import JSONPretty from 'react-json-pretty';
import 'react-json-pretty/themes/monikai.css';
import axios from 'axios';
import './App.css';

const { Header, Content, Footer } = Layout;
const { Title, Paragraph } = Typography;
const { Dragger } = Upload;

function App() {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [error, setError] = useState(null);

  const handleUpload = async (file) => {
    setLoading(true);
    setError(null);
    setResult(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await axios.post('/api/convert', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        timeout: 60000, // 60秒超时
      });

      if (response.data.success) {
        setResult(response.data.data);
        message.success('PPT文件转换成功！');
      } else {
        throw new Error(response.data.error || '转换失败');
      }
    } catch (err) {
      console.error('上传错误:', err);
      const errorMessage = err.response?.data?.message || err.message || '转换失败，请重试';
      setError(errorMessage);
      message.error(errorMessage);
    } finally {
      setLoading(false);
    }

    return false; // 阻止默认上传行为
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
    message.success('JSON文件下载成功！');
  };

  const uploadProps = {
    name: 'file',
    multiple: false,
    accept: '.ppt,.pptx',
    beforeUpload: handleUpload,
    showUploadList: false,
  };

  return (
    <Layout className="layout">
      <Header className="header">
        <div className="header-content">
          <FileTextOutlined className="logo" />
          <Title level={3} style={{ color: 'white', margin: 0 }}>
            PPT转JSON工具
          </Title>
        </div>
      </Header>

      <Content className="content">
        <div className="container">
          <Card className="main-card">
            <div className="upload-section">
              <Title level={4}>上传PPT文件</Title>
              <Paragraph>
                支持PPT和PPTX格式，最大文件大小50MB。转换后的JSON将包含幻灯片的所有元素信息。
              </Paragraph>

              <Dragger {...uploadProps} className="upload-dragger">
                <p className="ant-upload-drag-icon">
                  <UploadOutlined />
                </p>
                <p className="ant-upload-text">点击或拖拽PPT文件到此区域上传</p>
                <p className="ant-upload-hint">
                  支持 .ppt 和 .pptx 格式文件
                </p>
              </Dragger>

              {loading && (
                <div className="loading-section">
                  <Spin size="large" />
                  <p>正在转换PPT文件，请稍候...</p>
                </div>
              )}

              {error && (
                <Alert
                  message="转换失败"
                  description={error}
                  type="error"
                  showIcon
                  className="error-alert"
                />
              )}
            </div>

            {result && (
              <div className="result-section">
                <div className="result-header">
                  <Title level={4}>转换结果</Title>
                  <Button
                    type="primary"
                    icon={<DownloadOutlined />}
                    onClick={downloadJSON}
                  >
                    下载JSON
                  </Button>
                </div>

                <div className="json-container">
                  <JSONPretty
                    data={result}
                    theme="monikai"
                    style={{
                      backgroundColor: '#272822',
                      padding: '16px',
                      borderRadius: '8px',
                      fontSize: '14px',
                      maxHeight: '600px',
                      overflow: 'auto'
                    }}
                  />
                </div>

                <div className="result-summary">
                  <Alert
                    message="转换完成"
                    description={`成功转换 ${result.slides?.length || 0} 张幻灯片，包含 ${result.slides?.reduce((total, slide) => total + (slide.elements?.length || 0), 0) || 0} 个元素`}
                    type="success"
                    showIcon
                  />
                </div>
              </div>
            )}
          </Card>
        </div>
      </Content>

      <Footer className="footer">
        <div className="footer-content">
          <p>PPT转JSON工具 ©2024 - 将PowerPoint文件转换为结构化JSON数据</p>
          <p>支持文字、图片、形状、表格、图表等多种元素类型</p>
        </div>
      </Footer>
    </Layout>
  );
}

export default App; 