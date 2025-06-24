const PPTXParser = require('./utils/pptxParser');
const path = require('path');

async function testParser() {
  try {
    console.log('开始测试PPTX解析器...');
    
    // 检查uploads目录中是否有文件
    const fs = require('fs-extra');
    const uploadsDir = path.join(__dirname, 'uploads');
    
    if (!await fs.pathExists(uploadsDir)) {
      console.log('uploads目录不存在');
      return;
    }
    
    const files = await fs.readdir(uploadsDir);
    console.log('uploads目录中的文件:', files);
    
    if (files.length === 0) {
      console.log('uploads目录为空，请先上传一个PPTX文件');
      return;
    }
    
    // 使用第一个PPTX文件进行测试
    const testFile = files.find(file => file.endsWith('.pptx') || file.endsWith('.ppt'));
    if (!testFile) {
      console.log('未找到PPTX文件');
      return;
    }
    
    const filePath = path.join(uploadsDir, testFile);
    console.log(`测试文件: ${filePath}`);
    
    const parser = new PPTXParser();
    const result = await parser.parse(filePath);
    
    console.log('解析结果:');
    console.log('主题颜色:', Object.keys(result.themeColors || {}));
    console.log('幻灯片数量:', result.slides.length);
    
    if (result.slides.length > 0) {
      console.log('第一张幻灯片元素数量:', result.slides[0].elements.length);
      if (result.slides[0].elements.length > 0) {
        console.log('第一个元素:', result.slides[0].elements[0]);
      }
    }
    
    console.log('测试完成');
    
  } catch (error) {
    console.error('测试失败:', error);
  }
}

testParser(); 