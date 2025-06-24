const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs-extra');
const { v4: uuidv4 } = require('uuid');
const PPTXParser = require('./utils/pptxParser');

const app = express();
const PORT = process.env.PORT || 5001;

// 中间件
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'client/build')));

// 配置文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'uploads');
    fs.ensureDirSync(uploadDir);
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueName = `${uuidv4()}-${file.originalname}`;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'application/vnd.ms-powerpoint'
    ];
    
    if (allowedTypes.includes(file.mimetype) || file.originalname.endsWith('.pptx') || file.originalname.endsWith('.ppt')) {
      cb(null, true);
    } else {
      cb(new Error('只支持PPT和PPTX文件格式'), false);
    }
  },
  limits: {
    fileSize: 50 * 1024 * 1024 // 50MB限制
  }
});

// API路由
app.post('/api/convert', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '请上传PPT文件' });
    }

    const filePath = req.file.path;
    const parser = new PPTXParser();
    
    // 解析PPT文件
    const result = await parser.parse(filePath);
    
    // 清理上传的文件
    await fs.remove(filePath);
    
    res.json({
      success: true,
      data: result
    });
    
  } catch (error) {
    console.error('转换错误:', error);
    
    // 清理上传的文件
    if (req.file) {
      await fs.remove(req.file.path).catch(() => {});
    }
    
    res.status(500).json({
      error: '转换失败',
      message: error.message
    });
  }
});

// 健康检查
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'PPT转JSON服务运行正常' });
});

// 服务React应用
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'client/build', 'index.html'));
});

// 错误处理中间件
app.use((error, req, res, next) => {
  console.error('服务器错误:', error);
  res.status(500).json({
    error: '服务器内部错误',
    message: error.message
  });
});

app.listen(PORT, () => {
  console.log(`服务器运行在端口 ${PORT}`);
  console.log(`访问 http://localhost:${PORT} 使用Web界面`);
  console.log(`API接口: http://localhost:${PORT}/api/convert`);
}); 