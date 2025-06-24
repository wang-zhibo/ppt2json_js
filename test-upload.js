const fs = require('fs-extra');
const path = require('path');
const FormData = require('form-data');
const fetch = require('node-fetch');

async function testUpload() {
  try {
    console.log('=== 测试上传功能 ===\n');
    
    // 检查是否有测试文件
    const uploadsDir = path.join(__dirname, 'uploads');
    const files = await fs.readdir(uploadsDir);
    const pptxFiles = files.filter(file => file.endsWith('.pptx') || file.endsWith('.ppt'));
    
    if (pptxFiles.length === 0) {
      console.log('❌ 没有找到测试文件');
      console.log('请将PPTX文件放到uploads目录中');
      return;
    }
    
    const testFile = pptxFiles[0];
    const filePath = path.join(uploadsDir, testFile);
    console.log(`📄 测试文件: ${testFile}`);
    
    // 创建FormData
    const formData = new FormData();
    formData.append('file', fs.createReadStream(filePath));
    
    console.log('\n🔄 开始上传测试...');
    
    // 发送请求
    const response = await fetch('http://localhost:5001/api/convert', {
      method: 'POST',
      body: formData
    });
    
    console.log(`📊 响应状态: ${response.status}`);
    
    if (response.ok) {
      const data = await response.json();
      console.log('✅ 上传成功！');
      console.log('\n📋 解析结果:');
      console.log(`  主题颜色: ${Object.keys(data.data.themeColors || {}).length} 种`);
      console.log(`  幻灯片数量: ${data.data.slides.length}`);
      console.log(`  尺寸: ${data.data.width} x ${data.data.height} (EMU)`);
      
      if (data.data.slides.length > 0) {
        console.log('\n📄 幻灯片详情:');
        data.data.slides.forEach((slide, index) => {
          console.log(`  幻灯片 ${index + 1}: ${slide.elements.length} 个元素`);
          if (slide.error) {
            console.log(`    ❌ 错误: ${slide.error}`);
          }
        });
      } else {
        console.log('\n⚠️  警告: 未解析到任何幻灯片内容');
      }
      
      // 保存结果到文件
      const outputPath = path.join(__dirname, 'upload-test-result.json');
      await fs.writeJson(outputPath, data.data, { spaces: 2 });
      console.log(`\n💾 结果已保存到: ${outputPath}`);
      
    } else {
      const errorData = await response.json();
      console.log('❌ 上传失败');
      console.log('错误信息:', errorData);
    }
    
  } catch (error) {
    console.error('❌ 测试失败:', error.message);
    console.log('\n可能的原因:');
    console.log('1. 后端服务器未启动');
    console.log('2. 端口5001被占用');
    console.log('3. 网络连接问题');
  }
}

// 检查依赖
async function checkDependencies() {
  try {
    require('node-fetch');
    require('form-data');
    console.log('✅ 依赖检查通过');
    await testUpload();
  } catch (error) {
    console.log('❌ 缺少依赖，请安装:');
    console.log('npm install node-fetch form-data');
  }
}

checkDependencies(); 