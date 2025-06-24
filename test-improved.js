const PPTXParser = require('./utils/pptxParser');
const path = require('path');
const fs = require('fs-extra');

async function testImprovedParser() {
  try {
    console.log('=== 测试改进后的PPTX解析器 ===\n');
    
    // 检查uploads目录
    const uploadsDir = path.join(__dirname, 'uploads');
    if (!await fs.pathExists(uploadsDir)) {
      console.log('uploads目录不存在，创建中...');
      await fs.ensureDir(uploadsDir);
    }
    
    const files = await fs.readdir(uploadsDir);
    console.log('uploads目录中的文件:', files);
    
    if (files.length === 0) {
      console.log('\n📁 uploads目录为空');
      console.log('请上传一个PPTX文件到uploads目录，然后重新运行此测试');
      console.log('或者将PPTX文件路径作为参数传入: node test-improved.js <文件路径>');
      return;
    }
    
    // 查找PPTX文件
    const pptxFiles = files.filter(file => file.endsWith('.pptx') || file.endsWith('.ppt'));
    if (pptxFiles.length === 0) {
      console.log('\n❌ 未找到PPTX文件');
      return;
    }
    
    const testFile = pptxFiles[0];
    const filePath = path.join(uploadsDir, testFile);
    console.log(`\n📄 测试文件: ${testFile}`);
    console.log(`📂 完整路径: ${filePath}`);
    
    // 检查文件大小
    const stats = await fs.stat(filePath);
    console.log(`📊 文件大小: ${(stats.size / 1024 / 1024).toFixed(2)} MB`);
    
    console.log('\n🔄 开始解析...\n');
    
    const parser = new PPTXParser();
    const result = await parser.parse(filePath);
    
    console.log('\n✅ 解析完成！');
    console.log('\n📊 解析结果摘要:');
    console.log(`  主题颜色: ${Object.keys(result.themeColors || {}).length} 种`);
    console.log(`  幻灯片数量: ${result.slides.length}`);
    console.log(`  尺寸: ${result.width} x ${result.height} (EMU)`);
    
    if (result.slides.length > 0) {
      console.log('\n📋 幻灯片详情:');
      result.slides.forEach((slide, index) => {
        console.log(`  幻灯片 ${index + 1}:`);
        console.log(`    元素数量: ${slide.elements.length}`);
        if (slide.error) {
          console.log(`    ❌ 错误: ${slide.error}`);
        }
        if (slide.elements.length > 0) {
          const textElements = slide.elements.filter(el => el.type === 'text' && el.content);
          const shapeElements = slide.elements.filter(el => el.type === 'shape');
          const imageElements = slide.elements.filter(el => el.type === 'image');
          
          console.log(`    文本元素: ${textElements.length}`);
          console.log(`    形状元素: ${shapeElements.length}`);
          console.log(`    图片元素: ${imageElements.length}`);
          
          if (textElements.length > 0) {
            console.log(`    文本内容示例: "${textElements[0].content.substring(0, 50)}..."`);
          }
        }
      });
    } else {
      console.log('\n⚠️  警告: 未解析到任何幻灯片内容');
      console.log('可能的原因:');
      console.log('  1. 文件格式不兼容');
      console.log('  2. 文件损坏');
      console.log('  3. 解析器需要进一步调整');
    }
    
    // 保存结果到文件
    const outputPath = path.join(__dirname, 'test-result.json');
    await fs.writeJson(outputPath, result, { spaces: 2 });
    console.log(`\n💾 完整结果已保存到: ${outputPath}`);
    
  } catch (error) {
    console.error('\n❌ 测试失败:', error.message);
    console.error('详细错误:', error);
  }
}

// 如果有命令行参数，使用指定的文件
const args = process.argv.slice(2);
if (args.length > 0) {
  const filePath = args[0];
  if (fs.pathExistsSync(filePath)) {
    console.log(`使用指定文件: ${filePath}`);
    // 复制到uploads目录
    const uploadsDir = path.join(__dirname, 'uploads');
    fs.ensureDirSync(uploadsDir);
    const fileName = path.basename(filePath);
    const destPath = path.join(uploadsDir, fileName);
    fs.copyFileSync(filePath, destPath);
    console.log(`文件已复制到: ${destPath}`);
  } else {
    console.error(`文件不存在: ${filePath}`);
    process.exit(1);
  }
}

testImprovedParser(); 