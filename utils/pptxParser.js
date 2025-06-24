const fs = require('fs-extra');
const path = require('path');
const xml2js = require('xml2js');
const AdmZip = require('adm-zip');

class PPTXParser {
  constructor() {
    this.parser = new xml2js.Parser();
  }

  async parse(filePath) {
    try {
      // 解压PPTX文件
      const zip = new AdmZip(filePath);
      const tempDir = path.join(__dirname, '../temp', Date.now().toString());
      await fs.ensureDir(tempDir);
      
      zip.extractAllTo(tempDir, true);

      // 解析演示文稿结构
      const presentation = await this.parsePresentation(tempDir);
      
      // 清理临时文件
      await fs.remove(tempDir);
      
      return presentation;
    } catch (error) {
      throw new Error(`解析PPT文件失败: ${error.message}`);
    }
  }

  async parsePresentation(tempDir) {
    try {
      const presentationPath = path.join(tempDir, 'ppt/presentation.xml');
      
      if (!await fs.pathExists(presentationPath)) {
        console.error('presentation.xml文件不存在');
        throw new Error('presentation.xml文件不存在');
      }
      
      const presentationXml = await fs.readFile(presentationPath, 'utf8');
      const presentation = await this.parser.parseStringPromise(presentationXml);

      console.log('presentation.xml结构:', Object.keys(presentation));

      const result = {
        themeColors: await this.parseThemeColors(tempDir),
        size: await this.parseSlideSize(tempDir),
        width: 0,
        height: 0,
        slides: []
      };

      // 设置尺寸
      if (result.size) {
        result.width = result.size.cx || 9144000; // EMU单位
        result.height = result.size.cy || 6858000;
      }

      // 解析幻灯片 - 支持不同的XML结构
      console.log('开始解析幻灯片...');
      
      let slideIds = [];
      const sldIdLst = presentation['p:presentation']?.['p:sldIdLst']?.[0]?.['p:sldId'];
      
      console.log('完整的幻灯片ID列表 (sldIdLst):', JSON.stringify(sldIdLst, null, 2));

      if (sldIdLst && Array.isArray(sldIdLst)) {
        slideIds = sldIdLst;
      } else {
        console.log('警告: sldIdLst 不是一个数组或不存在, 将尝试备用方案.');
      }
            
      console.log(`找到 ${slideIds.length} 张幻灯片`);
      
      if (slideIds.length === 0) {
        console.log('未找到幻灯片ID列表，尝试其他方法...');
        // 尝试直接从slides目录读取
        const slidesDir = path.join(tempDir, 'ppt/slides');
        if (await fs.pathExists(slidesDir)) {
          const slideFiles = await fs.readdir(slidesDir);
          const pptxSlideFiles = slideFiles.filter(file => file.endsWith('.xml'));
          console.log(`从slides目录找到 ${pptxSlideFiles.length} 个幻灯片文件`);
          
          for (const slideFile of pptxSlideFiles) {
            try {
              const slidePath = `slides/${slideFile}`;
              console.log(`解析幻灯片文件: ${slidePath}`);
              const slide = await this.parseSlide(tempDir, slidePath);
              result.slides.push(slide);
            } catch (error) {
              console.error(`解析幻灯片 ${slideFile} 失败:`, error.message);
            }
          }
        }
      } else {
        for (const slideId of slideIds) {
          const slideRid = slideId?.$?.['r:id'];
          
          if (!slideRid) {
            console.error('无法从幻灯片ID对象中找到 r:id:', JSON.stringify(slideId));
            continue;
          }

          console.log(`解析幻灯片 ID: ${slideRid}`);
          
          const slidePath = await this.getSlidePath(tempDir, slideRid);
          console.log(`幻灯片路径: ${slidePath}`);
          
          if (slidePath) {
            try {
              const slide = await this.parseSlide(tempDir, slidePath);
              console.log(`幻灯片解析成功，包含 ${slide.elements.length} 个元素`);
              result.slides.push(slide);
            } catch (error) {
              console.error(`解析幻灯片失败: ${error.message}`);
              // 添加一个空的幻灯片以保持索引
              result.slides.push({
                fill: null,
                note: null,
                elements: [],
                layoutElements: [],
                error: error.message
              });
            }
          } else {
            console.error(`未找到幻灯片路径: ${slideRid}`);
          }
        }
      }

      console.log(`解析完成，总共 ${result.slides.length} 张幻灯片`);
      return result;
    } catch (error) {
      console.error('解析演示文稿失败:', error.message);
      throw error;
    }
  }

  async parseThemeColors(tempDir) {
    try {
      const themePath = path.join(tempDir, 'ppt/theme/theme1.xml');
      const themeXml = await fs.readFile(themePath, 'utf8');
      const theme = await this.parser.parseStringPromise(themeXml);
      
      const colorScheme = theme['a:theme']['a:themeElements'][0]['a:clrScheme'][0];
      const colors = {};
      
      ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].forEach(key => {
        if (colorScheme[`a:${key}`]) {
          const color = colorScheme[`a:${key}`][0];
          if (color['a:srgbClr']) {
            colors[key] = color['a:srgbClr'][0].$.val;
          }
        }
      });
      
      return colors;
    } catch (error) {
      return {};
    }
  }

  async parseSlideSize(tempDir) {
    try {
      const presentationPath = path.join(tempDir, 'ppt/presentation.xml');
      const presentationXml = await fs.readFile(presentationPath, 'utf8');
      const presentation = await this.parser.parseStringPromise(presentationXml);
      
      const sldSz = presentation['p:presentation']['p:sldSz'][0].$;
      return {
        cx: parseInt(sldSz.cx),
        cy: parseInt(sldSz.cy),
        type: sldSz.type
      };
    } catch (error) {
      return null;
    }
  }

  async getSlidePath(tempDir, slideRid) {
    try {
      console.log(`查找幻灯片路径，ID: ${slideRid}`);
      const relsPath = path.join(tempDir, 'ppt/_rels/presentation.xml.rels');
      
      // 检查关系文件是否存在
      if (!await fs.pathExists(relsPath)) {
        console.error('关系文件不存在:', relsPath);
        return null;
      }
      
      const relsXml = await fs.readFile(relsPath, 'utf8');
      const rels = await this.parser.parseStringPromise(relsXml);
      
      console.log('关系文件 (rels) 完整内容:', JSON.stringify(rels, null, 2));
      
      if (!rels.Relationships || !rels.Relationships.Relationship) {
        console.error('关系文件中没有找到Relationship元素');
        return null;
      }
      
      const relationships = Array.isArray(rels.Relationships.Relationship) 
        ? rels.Relationships.Relationship 
        : [rels.Relationships.Relationship];
      
      console.log(`找到 ${relationships.length} 个关系`);
      
      for (const rel of relationships) {
        const relId = rel.$.Id;
        const target = rel.$.Target;
        console.log(`检查关系: Id=${relId}, Target=${target}`);
        if (relId === slideRid) {
          const slidePath = target.replace('../', '');
          console.log(`找到匹配的幻灯片路径: ${slidePath}`);
          return slidePath;
        }
      }
      
      console.error(`未找到ID为 ${slideRid} 的幻灯片关系`);
      return null;
    } catch (error) {
      console.error('解析关系文件时出错:', error.message);
      return null;
    }
  }

  async parseSlide(tempDir, slidePath) {
    try {
      console.log(`开始解析幻灯片文件: ${slidePath}`);
      const fullSlidePath = path.join(tempDir, 'ppt', slidePath);
      
      // 检查文件是否存在
      if (!await fs.pathExists(fullSlidePath)) {
        console.error(`幻灯片文件不存在: ${fullSlidePath}`);
        throw new Error(`幻灯片文件不存在: ${slidePath}`);
      }
      
      const slideXml = await fs.readFile(fullSlidePath, 'utf8');
      const slide = await this.parser.parseStringPromise(slideXml);

      console.log('幻灯片XML结构:', Object.keys(slide));
      
      // 支持不同的XML结构
      let slideData = null;
      if (slide['p:sld'] && slide['p:sld']['p:cSld']) {
        slideData = slide['p:sld']['p:cSld'][0];
      } else if (slide.sld && slide.sld.cSld) {
        slideData = slide.sld.cSld[0];
      } else {
        console.error('幻灯片XML结构不正确');
        throw new Error('幻灯片XML结构不正确');
      }
      
      console.log('幻灯片数据结构:', Object.keys(slideData));
      
      const result = {
        fill: await this.parseFill(slideData['p:bg'] || slideData.bg),
        note: await this.parseNote(tempDir, slidePath),
        elements: [],
        layoutElements: []
      };

      // 解析页面元素 - 支持不同的命名空间
      let spTree = null;
      if (slideData['p:spTree']) {
        spTree = slideData['p:spTree'][0];
      } else if (slideData.spTree) {
        spTree = slideData.spTree[0];
      }
      
      if (spTree) {
        console.log('找到spTree，开始解析元素...');
        result.elements = await this.parseElements(spTree);
        console.log(`解析完成，找到 ${result.elements.length} 个元素`);
      } else {
        console.log('未找到spTree元素');
      }

      return result;
    } catch (error) {
      console.error(`解析幻灯片失败: ${error.message}`);
      throw error;
    }
  }

  async parseFill(bgElement) {
    if (!bgElement) return null;
    
    const bgPr = bgElement[0]['p:bgPr'];
    if (!bgPr) return null;

    const fill = {};
    
    if (bgPr[0]['a:solidFill']) {
      const solidFill = bgPr[0]['a:solidFill'][0];
      if (solidFill['a:srgbClr']) {
        fill.type = 'color';
        fill.color = solidFill['a:srgbClr'][0].$.val;
      }
    } else if (bgPr[0]['a:gradFill']) {
      fill.type = 'gradient';
      fill.gradient = this.parseGradient(bgPr[0]['a:gradFill'][0]);
    } else if (bgPr[0]['a:blipFill']) {
      fill.type = 'image';
      fill.image = this.parseImageFill(bgPr[0]['a:blipFill'][0]);
    }

    return fill;
  }

  parseGradient(gradFill) {
    const gradient = {};
    
    if (gradFill['a:gsLst']) {
      const stops = gradFill['a:gsLst'][0]['a:gs'];
      gradient.stops = stops.map(stop => ({
        position: parseFloat(stop.$.pos) / 1000,
        color: stop['a:srgbClr'] ? stop['a:srgbClr'][0].$.val : null
      }));
    }

    if (gradFill['a:lin']) {
      gradient.type = 'linear';
      const lin = gradFill['a:lin'][0].$;
      gradient.angle = Math.atan2(parseInt(lin.ang), 60000) * 180 / Math.PI;
    }

    return gradient;
  }

  parseImageFill(blipFill) {
    const image = {};
    
    if (blipFill['a:blip']) {
      const blip = blipFill['a:blip'][0];
      image.embed = blip.$.rId;
    }

    if (blipFill['a:stretch']) {
      image.stretch = 'stretch';
    } else if (blipFill['a:tile']) {
      image.stretch = 'tile';
    }

    return image;
  }

  async parseNote(tempDir, slidePath) {
    try {
      const slideRelsPath = path.join(tempDir, 'ppt/slides/_rels', path.basename(slidePath) + '.rels');
      const slideRelsXml = await fs.readFile(slideRelsPath, 'utf8');
      const slideRels = await this.parser.parseStringPromise(slideRelsXml);
      
      const relationships = slideRels.Relationships.Relationship;
      for (const rel of relationships) {
        if (rel.$.Type.includes('notesSlide')) {
          const notePath = path.join(tempDir, 'ppt', rel.$.Target);
          const noteXml = await fs.readFile(notePath, 'utf8');
          const note = await this.parser.parseStringPromise(noteXml);
          
          if (note['p:notes']['p:cSld'][0]['p:spTree']) {
            const spTree = note['p:notes']['p:cSld'][0]['p:spTree'][0];
            return this.extractTextFromSpTree(spTree);
          }
        }
      }
    } catch (error) {
      return '';
    }
    return '';
  }

  async parseElements(spTree) {
    const elements = [];
    
    console.log('spTree包含的元素类型:', Object.keys(spTree));
    
    // 支持不同的命名空间
    const elementTypes = [
      { key: 'p:sp', name: '形状元素' },
      { key: 'sp', name: '形状元素(无命名空间)' },
      { key: 'p:pic', name: '图片元素' },
      { key: 'pic', name: '图片元素(无命名空间)' },
      { key: 'p:graphicFrame', name: '图形框架元素' },
      { key: 'graphicFrame', name: '图形框架元素(无命名空间)' },
      { key: 'p:grpSp', name: '组合元素' },
      { key: 'grpSp', name: '组合元素(无命名空间)' }
    ];
    
    for (const elementType of elementTypes) {
      if (spTree[elementType.key]) {
        console.log(`找到 ${spTree[elementType.key].length} 个${elementType.name}`);
        const items = spTree[elementType.key];
        
        for (const item of items) {
          let element = null;
          
          try {
            if (elementType.key.includes('sp')) {
              element = await this.parseShape(item);
            } else if (elementType.key.includes('pic')) {
              element = await this.parsePicture(item);
            } else if (elementType.key.includes('graphicFrame')) {
              element = await this.parseGraphicFrame(item);
            } else if (elementType.key.includes('grpSp')) {
              element = await this.parseGroup(item);
            }
            
            if (element) {
              elements.push(element);
            }
          } catch (error) {
            console.error(`解析${elementType.name}失败:`, error.message);
          }
        }
      }
    }

    console.log(`总共解析出 ${elements.length} 个元素`);
    return elements;
  }

  async parseShape(sp) {
    console.log('解析形状元素...');
    
    // 支持不同的命名空间
    const spPr = sp['p:spPr'] ? sp['p:spPr'][0] : (sp.spPr ? sp.spPr[0] : null);
    const txBody = sp['p:txBody'] || sp.txBody;
    
    if (!spPr) {
      console.log('未找到形状属性，跳过此元素');
      return null;
    }
    
    const element = {
      type: 'text',
      ...this.parseShapeProperties(spPr),
      content: txBody ? this.extractTextFromTxBody(txBody[0] || txBody) : '',
      vAlign: this.getVerticalAlignment(txBody),
      isVertical: this.isVerticalText(txBody)
    };

    console.log('形状元素内容:', element.content);
    console.log('形状元素类型:', element.type);

    // 检查是否为形状 - 支持不同的命名空间
    const nvSpPr = sp['p:nvSpPr'] || sp.nvSpPr;
    if (nvSpPr) {
      const nvSpPrData = Array.isArray(nvSpPr) ? nvSpPr[0] : nvSpPr;
      const cNvPr = nvSpPrData['p:cNvPr'] || nvSpPrData.cNvPr;
      
      if (cNvPr) {
        const cNvPrData = Array.isArray(cNvPr) ? cNvPr[0] : cNvPr;
        if (cNvPrData.$.name && !cNvPrData.$.name.includes('TextBox')) {
          element.type = 'shape';
          element.shapeType = this.getShapeType(spPr);
          element.name = cNvPrData.$.name;
          console.log('检测到形状:', element.name, '类型:', element.shapeType);
        }
      }
    }

    return element;
  }

  parseShapeProperties(spPr) {
    const props = {};
    
    if (spPr['a:xfrm']) {
      const xfrm = spPr['a:xfrm'][0];
      if (xfrm['a:off']) {
        props.left = parseInt(xfrm['a:off'][0].$.x);
        props.top = parseInt(xfrm['a:off'][0].$.y);
      }
      if (xfrm['a:ext']) {
        props.width = parseInt(xfrm['a:ext'][0].$.cx);
        props.height = parseInt(xfrm['a:ext'][0].$.cy);
      }
      if (xfrm['a:rot']) {
        props.rotate = parseFloat(xfrm['a:rot'][0]) / 60000;
      }
      if (xfrm['a:flipH']) {
        props.isFlipH = true;
      }
      if (xfrm['a:flipV']) {
        props.isFlipV = true;
      }
    }

    // 边框
    if (spPr['a:ln']) {
      const ln = spPr['a:ln'][0];
      props.borderColor = ln['a:solidFill'] ? this.getColor(ln['a:solidFill'][0]) : null;
      props.borderWidth = ln.$.w ? parseInt(ln.$.w) : 0;
      props.borderType = ln['a:prstDash'] ? ln['a:prstDash'][0].$.val : 'solid';
      if (ln['a:custDash']) {
        props.borderStrokeDasharray = this.parseDashArray(ln['a:custDash'][0]);
      }
    }

    // 填充
    if (spPr['a:solidFill']) {
      props.fill = { type: 'color', color: this.getColor(spPr['a:solidFill'][0]) };
    } else if (spPr['a:gradFill']) {
      props.fill = { type: 'gradient', gradient: this.parseGradient(spPr['a:gradFill'][0]) };
    } else if (spPr['a:blipFill']) {
      props.fill = { type: 'image', image: this.parseImageFill(spPr['a:blipFill'][0]) };
    }

    // 阴影
    if (spPr['a:effectLst']) {
      props.shadow = this.parseShadow(spPr['a:effectLst'][0]);
    }

    return props;
  }

  getColor(solidFill) {
    if (solidFill['a:srgbClr']) {
      return solidFill['a:srgbClr'][0].$.val;
    } else if (solidFill['a:schemeClr']) {
      return solidFill['a:schemeClr'][0].$.val;
    }
    return null;
  }

  parseDashArray(custDash) {
    if (custDash['a:dsLst']) {
      return custDash['a:dsLst'][0]['a:ds'].map(ds => parseInt(ds.$.d));
    }
    return [];
  }

  parseShadow(effectLst) {
    const shadow = {};
    
    if (effectLst['a:outerShdw']) {
      const outerShdw = effectLst['a:outerShdw'][0];
      shadow.type = 'outer';
      shadow.color = outerShdw['a:srgbClr'] ? outerShdw['a:srgbClr'][0].$.val : null;
      shadow.blur = outerShdw.$.blurRad ? parseInt(outerShdw.$.blurRad) : 0;
      shadow.distance = outerShdw.$.dist ? parseInt(outerShdw.$.dist) : 0;
      shadow.angle = outerShdw.$.dir ? parseInt(outerShdw.$.dir) : 0;
    }

    return shadow;
  }

  extractTextFromTxBody(txBody) {
    let text = '';
    
    console.log('开始提取文本，txBody结构:', Object.keys(txBody));
    
    // 支持不同的命名空间
    const paragraphs = txBody['a:p'] || txBody.p || [];
    
    if (paragraphs.length > 0) {
      console.log(`找到 ${paragraphs.length} 个段落`);
      for (const p of paragraphs) {
        const runs = p['a:r'] || p.r || [];
        if (runs.length > 0) {
          console.log(`段落包含 ${runs.length} 个文本运行`);
          for (const r of runs) {
            const textElements = r['a:t'] || r.t || [];
            if (textElements.length > 0) {
              const textContent = textElements[0];
              console.log('提取到文本:', textContent);
              text += textContent;
            }
          }
        }
        text += '\n';
      }
    }
    
    console.log('最终提取的文本:', text);
    return text.trim();
  }

  getVerticalAlignment(txBody) {
    if (!txBody) return 'top';
    
    const txBodyPr = txBody[0]['a:txBodyPr'];
    if (txBodyPr && txBodyPr[0]['a:anchor']) {
      return txBodyPr[0]['a:anchor'][0];
    }
    
    return 'top';
  }

  isVerticalText(txBody) {
    if (!txBody) return false;
    
    const txBodyPr = txBody[0]['a:txBodyPr'];
    if (txBodyPr && txBodyPr[0]['a:vert']) {
      return txBodyPr[0]['a:vert'][0] === 'vert';
    }
    
    return false;
  }

  getShapeType(spPr) {
    if (spPr['a:prstGeom']) {
      return spPr['a:prstGeom'][0].$.prst;
    }
    return 'rect';
  }

  async parsePicture(pic) {
    const blipFill = pic['p:blipFill'];
    if (!blipFill) return null;

    const spPr = pic['p:spPr'][0];
    const element = {
      type: 'image',
      ...this.parseShapeProperties(spPr)
    };

    if (blipFill[0]['a:blip']) {
      const blip = blipFill[0]['a:blip'][0];
      element.src = blip.$.rId;
    }

    if (blipFill[0]['a:srcRect']) {
      element.rect = blipFill[0]['a:srcRect'][0].$;
    }

    return element;
  }

  async parseGraphicFrame(frame) {
    const xfrm = frame['p:xfrm'][0];
    const element = {
      type: 'unknown',
      left: parseInt(xfrm['a:off'][0].$.x),
      top: parseInt(xfrm['a:off'][0].$.y),
      width: parseInt(xfrm['a:ext'][0].$.cx),
      height: parseInt(xfrm['a:ext'][0].$.cy)
    };

    // 检查图表类型
    if (frame['a:graphic'] && frame['a:graphic'][0]['a:graphicData']) {
      const graphicData = frame['a:graphic'][0]['a:graphicData'][0];
      
      if (graphicData['c:chart']) {
        element.type = 'chart';
        element.chartData = this.parseChartData(graphicData['c:chart'][0]);
      } else if (graphicData['dgm:relIds']) {
        element.type = 'diagram';
        element.diagramData = this.parseDiagramData(graphicData);
      } else if (graphicData['a:tbl']) {
        element.type = 'table';
        element.tableData = this.parseTableData(graphicData['a:tbl'][0]);
      }
    }

    return element;
  }

  parseChartData(chart) {
    const chartData = {
      chartType: 'unknown',
      data: [],
      colors: []
    };

    if (chart['c:chartType']) {
      chartData.chartType = chart['c:chartType'][0].$.val;
    }

    // 解析图表数据
    if (chart['c:ser']) {
      for (const ser of chart['c:ser']) {
        const series = {};
        if (ser['c:val']) {
          series.values = this.parseChartValues(ser['c:val'][0]);
        }
        if (ser['c:cat']) {
          series.categories = this.parseChartValues(ser['c:cat'][0]);
        }
        chartData.data.push(series);
      }
    }

    return chartData;
  }

  parseChartValues(val) {
    if (val['c:numRef'] && val['c:numRef'][0]['c:numCache']) {
      const numCache = val['c:numRef'][0]['c:numCache'][0];
      if (numCache['c:pt']) {
        return numCache['c:pt'].map(pt => parseFloat(pt['c:v'][0]));
      }
    }
    return [];
  }

  parseDiagramData(graphicData) {
    return {
      type: 'diagram',
      elements: []
    };
  }

  parseTableData(table) {
    const tableData = {
      data: [],
      rowHeights: [],
      colWidths: []
    };

    if (table['a:tblGrid']) {
      const grid = table['a:tblGrid'][0];
      if (grid['a:gridCol']) {
        tableData.colWidths = grid['a:gridCol'].map(col => parseInt(col.$.w));
      }
    }

    if (table['a:tr']) {
      for (const row of table['a:tr']) {
        const rowData = [];
        if (row['a:tc']) {
          for (const cell of row['a:tc']) {
            const cellText = this.extractTextFromSpTree(cell);
            rowData.push(cellText);
          }
        }
        tableData.data.push(rowData);
      }
    }

    return tableData;
  }

  async parseGroup(group) {
    const grpSpPr = group['p:grpSpPr'][0];
    const element = {
      type: 'group',
      ...this.parseShapeProperties(grpSpPr),
      elements: []
    };

    if (group['p:spTree']) {
      element.elements = await this.parseElements(group['p:spTree'][0]);
    }

    return element;
  }

  extractTextFromSpTree(spTree) {
    let text = '';
    
    if (spTree['p:txBody']) {
      text = this.extractTextFromTxBody(spTree['p:txBody'][0]);
    }
    
    return text;
  }
}

module.exports = PPTXParser; 