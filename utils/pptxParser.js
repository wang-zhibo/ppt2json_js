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
    const presentationPath = path.join(tempDir, 'ppt/presentation.xml');
    const presentationXml = await fs.readFile(presentationPath, 'utf8');
    const presentation = await this.parser.parseStringPromise(presentationXml);

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

    // 解析幻灯片
    const slideIds = presentation['p:presentation']['p:sldIdLst'][0]['p:sldId'];
    for (const slideId of slideIds) {
      const slideRid = slideId.$.rId;
      const slidePath = await this.getSlidePath(tempDir, slideRid);
      if (slidePath) {
        const slide = await this.parseSlide(tempDir, slidePath);
        result.slides.push(slide);
      }
    }

    return result;
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
      const relsPath = path.join(tempDir, 'ppt/_rels/presentation.xml.rels');
      const relsXml = await fs.readFile(relsPath, 'utf8');
      const rels = await this.parser.parseStringPromise(relsXml);
      
      const relationships = rels.Relationships.Relationship;
      for (const rel of relationships) {
        if (rel.$.Id === slideRid) {
          return rel.$.Target.replace('../', '');
        }
      }
    } catch (error) {
      return null;
    }
    return null;
  }

  async parseSlide(tempDir, slidePath) {
    const fullSlidePath = path.join(tempDir, 'ppt', slidePath);
    const slideXml = await fs.readFile(fullSlidePath, 'utf8');
    const slide = await this.parser.parseStringPromise(slideXml);

    const slideData = slide['p:sld']['p:cSld'][0];
    
    const result = {
      fill: await this.parseFill(slideData['p:bg']),
      note: await this.parseNote(tempDir, slidePath),
      elements: [],
      layoutElements: []
    };

    // 解析页面元素
    if (slideData['p:spTree']) {
      const spTree = slideData['p:spTree'][0];
      result.elements = await this.parseElements(spTree);
    }

    return result;
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
    
    if (spTree['p:sp']) {
      for (const sp of spTree['p:sp']) {
        const element = await this.parseShape(sp);
        if (element) elements.push(element);
      }
    }

    if (spTree['p:pic']) {
      for (const pic of spTree['p:pic']) {
        const element = await this.parsePicture(pic);
        if (element) elements.push(element);
      }
    }

    if (spTree['p:graphicFrame']) {
      for (const frame of spTree['p:graphicFrame']) {
        const element = await this.parseGraphicFrame(frame);
        if (element) elements.push(element);
      }
    }

    if (spTree['p:grpSp']) {
      for (const group of spTree['p:grpSp']) {
        const element = await this.parseGroup(group);
        if (element) elements.push(element);
      }
    }

    return elements;
  }

  async parseShape(sp) {
    const spPr = sp['p:spPr'][0];
    const txBody = sp['p:txBody'];
    
    const element = {
      type: 'text',
      ...this.parseShapeProperties(spPr),
      content: txBody ? this.extractTextFromTxBody(txBody[0]) : '',
      vAlign: this.getVerticalAlignment(txBody),
      isVertical: this.isVerticalText(txBody)
    };

    // 检查是否为形状
    if (sp['p:nvSpPr'] && sp['p:nvSpPr'][0]['p:cNvPr']) {
      const cNvPr = sp['p:nvSpPr'][0]['p:cNvPr'][0];
      if (cNvPr.$.name && !cNvPr.$.name.includes('TextBox')) {
        element.type = 'shape';
        element.shapeType = this.getShapeType(spPr);
        element.name = cNvPr.$.name;
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
    
    if (txBody['a:p']) {
      for (const p of txBody['a:p']) {
        if (p['a:r']) {
          for (const r of p['a:r']) {
            if (r['a:t']) {
              text += r['a:t'][0];
            }
          }
        }
        text += '\n';
      }
    }
    
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