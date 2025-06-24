const fs = require('fs-extra');
const path = require('path');
const xml2js = require('xml2js');
const AdmZip = require('adm-zip');

async function checkPresentation() {
  try {
    console.log('检查presentation.xml结构...');
    
    // 解压PPTX文件
    const zip = new AdmZip('uploads/ragflow.pptx');
    const tempDir = './debug-temp';
    await fs.ensureDir(tempDir);
    zip.extractAllTo(tempDir, true);
    
    // 读取presentation.xml
    const presentationPath = path.join(tempDir, 'ppt/presentation.xml');
    const presentationXml = await fs.readFile(presentationPath, 'utf8');
    
    console.log('presentation.xml内容:');
    console.log(presentationXml);
    
    // 解析XML
    const parser = new xml2js.Parser();
    const presentation = await parser.parseStringPromise(presentationXml);
    
    console.log('\n解析后的结构:');
    console.log(JSON.stringify(presentation, null, 2));
    
    // 检查幻灯片ID列表
    console.log('\n检查幻灯片ID列表...');
    if (presentation['p:presentation']) {
      console.log('找到 p:presentation');
      if (presentation['p:presentation']['p:sldIdLst']) {
        console.log('找到 p:sldIdLst');
        const slideIds = presentation['p:presentation']['p:sldIdLst'][0]['p:sldId'];
        console.log(`找到 ${slideIds.length} 个幻灯片ID:`, slideIds);
      } else {
        console.log('未找到 p:sldIdLst');
      }
    } else {
      console.log('未找到 p:presentation');
    }
    
    // 清理临时文件
    await fs.remove(tempDir);
    
  } catch (error) {
    console.error('检查失败:', error);
  }
}

checkPresentation(); 