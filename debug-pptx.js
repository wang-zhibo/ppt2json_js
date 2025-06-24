const AdmZip = require('adm-zip');
const fs = require('fs-extra');
const path = require('path');

async function debugPPTX(filePath) {
  try {
    console.log(`调试PPTX文件: ${filePath}`);
    
    if (!await fs.pathExists(filePath)) {
      console.error('文件不存在');
      return;
    }
    
    // 解压PPTX文件
    const zip = new AdmZip(filePath);
    const tempDir = path.join(__dirname, 'debug-temp');
    await fs.ensureDir(tempDir);
    
    zip.extractAllTo(tempDir, true);
    
    console.log('PPTX文件结构:');
    await printDirectoryStructure(tempDir, '', 0);
    
    // 检查关键文件
    const keyFiles = [
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/theme/theme1.xml'
    ];
    
    for (const keyFile of keyFiles) {
      const fullPath = path.join(tempDir, keyFile);
      if (await fs.pathExists(fullPath)) {
        console.log(`✓ ${keyFile} 存在`);
        const content = await fs.readFile(fullPath, 'utf8');
        console.log(`  大小: ${content.length} 字符`);
      } else {
        console.log(`✗ ${keyFile} 不存在`);
      }
    }
    
    // 清理临时文件
    await fs.remove(tempDir);
    
  } catch (error) {
    console.error('调试失败:', error);
  }
}

async function printDirectoryStructure(dir, prefix, depth) {
  if (depth > 3) return; // 限制深度避免过多输出
  
  const items = await fs.readdir(dir);
  for (const item of items) {
    const fullPath = path.join(dir, item);
    const stat = await fs.stat(fullPath);
    const relativePath = path.relative(path.join(__dirname, 'debug-temp'), fullPath);
    
    if (stat.isDirectory()) {
      console.log(`${prefix}📁 ${relativePath}/`);
      await printDirectoryStructure(fullPath, prefix + '  ', depth + 1);
    } else {
      console.log(`${prefix}📄 ${relativePath} (${stat.size} bytes)`);
    }
  }
}

// 如果有命令行参数，使用指定的文件
const args = process.argv.slice(2);
if (args.length > 0) {
  debugPPTX(args[0]);
} else {
  console.log('使用方法: node debug-pptx.js <pptx文件路径>');
} 