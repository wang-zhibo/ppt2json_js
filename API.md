# PPT转JSON API文档

## 概述

PPT转JSON工具提供RESTful API接口，支持将PowerPoint文件转换为结构化的JSON格式。

## 基础信息

- **基础URL**: `http://localhost:5001`
- **内容类型**: `application/json`
- **文件上传**: `multipart/form-data`

## 端点

### 1. 转换PPT文件

将PPT/PPTX文件转换为JSON格式。

**请求**
```
POST /api/convert
```

**参数**
| 参数名 | 类型 | 必需 | 描述 |
|--------|------|------|------|
| file | File | 是 | PPT或PPTX文件 |

**请求示例**
```bash
curl -X POST \
  http://localhost:5001/api/convert \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@presentation.pptx'
```

**响应格式**
```json
{
  "success": true,
  "data": {
    "themeColors": {
      "dk1": "000000",
      "lt1": "FFFFFF",
      "dk2": "404040",
      "lt2": "BFBFBF",
      "accent1": "4472C4",
      "accent2": "ED7D31",
      "accent3": "A5A5A5",
      "accent4": "FFC000",
      "accent5": "5B9BD5",
      "accent6": "70AD47",
      "hlink": "0563C1",
      "folHlink": "954F72"
    },
    "size": {
      "cx": 9144000,
      "cy": 6858000,
      "type": "screen4x3"
    },
    "width": 9144000,
    "height": 6858000,
    "slides": [
      {
        "fill": {
          "type": "color",
          "color": "FFFFFF"
        },
        "note": "这是第一张幻灯片的备注",
        "elements": [
          {
            "type": "text",
            "left": 914400,
            "top": 685800,
            "width": 7315200,
            "height": 1371600,
            "borderColor": null,
            "borderWidth": 0,
            "borderType": "solid",
            "fill": {
              "type": "color",
              "color": "FFFFFF"
            },
            "content": "这是标题文本",
            "rotate": 0,
            "vAlign": "top",
            "isVertical": false
          }
        ],
        "layoutElements": []
      }
    ]
  }
}
```

**错误响应**
```json
{
  "error": "转换失败",
  "message": "文件格式不支持"
}
```

### 2. 健康检查

检查服务运行状态。

**请求**
```
GET /api/health
```

**响应**
```json
{
  "status": "ok",
  "message": "PPT转JSON服务运行正常"
}
```

## 数据格式说明

### 幻灯片数据结构

```json
{
  "themeColors": {},      // 主题色配置
  "size": {},            // 幻灯片尺寸
  "width": 9144000,      // 宽度（EMU单位）
  "height": 6858000,     // 高度（EMU单位）
  "slides": []           // 幻灯片数组
}
```

### 元素类型

#### 文字元素 (type: "text")
```json
{
  "type": "text",
  "left": 914400,        // 水平坐标
  "top": 685800,         // 垂直坐标
  "width": 7315200,      // 宽度
  "height": 1371600,     // 高度
  "borderColor": "#000000", // 边框颜色
  "borderWidth": 12700,  // 边框宽度
  "borderType": "solid", // 边框类型
  "borderStrokeDasharray": [], // 虚线样式
  "shadow": {},          // 阴影效果
  "fill": {},            // 填充样式
  "content": "文本内容", // 文本内容
  "rotate": 0,           // 旋转角度
  "vAlign": "top",       // 垂直对齐
  "isVertical": false    // 是否竖向文本
}
```

#### 图片元素 (type: "image")
```json
{
  "type": "image",
  "left": 914400,
  "top": 685800,
  "width": 7315200,
  "height": 1371600,
  "borderColor": "#000000",
  "borderWidth": 12700,
  "borderType": "solid",
  "geom": "rect",        // 裁剪形状
  "rect": {},            // 裁剪范围
  "src": "rId1",         // 图片引用ID
  "rotate": 0
}
```

#### 形状元素 (type: "shape")
```json
{
  "type": "shape",
  "left": 914400,
  "top": 685800,
  "width": 7315200,
  "height": 1371600,
  "borderColor": "#000000",
  "borderWidth": 12700,
  "borderType": "solid",
  "shadow": {},
  "fill": {},
  "content": "形状文本",
  "rotate": 0,
  "shapeType": "rect",   // 形状类型
  "path": "",            // 形状路径
  "name": "形状1"        // 元素名称
}
```

#### 表格元素 (type: "table")
```json
{
  "type": "table",
  "left": 914400,
  "top": 685800,
  "width": 7315200,
  "height": 1371600,
  "borders": {},         // 边框样式
  "data": [              // 表格数据
    ["单元格1", "单元格2"],
    ["单元格3", "单元格4"]
  ],
  "rowHeights": [1371600, 1371600], // 行高
  "colWidths": [3657600, 3657600]   // 列宽
}
```

#### 图表元素 (type: "chart")
```json
{
  "type": "chart",
  "left": 914400,
  "top": 685800,
  "width": 7315200,
  "height": 1371600,
  "data": [              // 图表数据
    {
      "values": [10, 20, 30],
      "categories": ["A", "B", "C"]
    }
  ],
  "colors": ["#4472C4", "#ED7D31"], // 图表颜色
  "chartType": "barChart",          // 图表类型
  "barDir": "col",                  // 柱状图方向
  "marker": true,                   // 数据标记
  "holeSize": 50,                   // 环形图尺寸
  "grouping": "clustered",          // 分组模式
  "style": {}                       // 图表样式
}
```

#### 组合元素 (type: "group")
```json
{
  "type": "group",
  "left": 914400,
  "top": 685800,
  "width": 7315200,
  "height": 1371600,
  "elements": []        // 子元素数组
}
```

### 填充样式

#### 颜色填充
```json
{
  "type": "color",
  "color": "#FFFFFF"
}
```

#### 渐变填充
```json
{
  "type": "gradient",
  "gradient": {
    "type": "linear",
    "angle": 45,
    "stops": [
      {
        "position": 0,
        "color": "#FFFFFF"
      },
      {
        "position": 1,
        "color": "#000000"
      }
    ]
  }
}
```

#### 图片填充
```json
{
  "type": "image",
  "image": {
    "embed": "rId1",
    "stretch": "stretch"
  }
}
```

## 错误代码

| 状态码 | 错误类型 | 描述 |
|--------|----------|------|
| 400 | Bad Request | 请求参数错误 |
| 413 | Payload Too Large | 文件过大 |
| 415 | Unsupported Media Type | 不支持的文件格式 |
| 500 | Internal Server Error | 服务器内部错误 |

## 使用限制

- 最大文件大小：50MB
- 支持格式：.ppt, .pptx
- 请求超时：60秒
- 并发请求：无限制（受服务器性能影响）

## 示例代码

### JavaScript (Fetch API)
```javascript
async function convertPPT(file) {
  const formData = new FormData();
  formData.append('file', file);

  try {
    const response = await fetch('/api/convert', {
      method: 'POST',
      body: formData
    });

    const result = await response.json();
    
    if (result.success) {
      console.log('转换成功:', result.data);
      return result.data;
    } else {
      throw new Error(result.message);
    }
  } catch (error) {
    console.error('转换失败:', error);
    throw error;
  }
}
```

### Python (requests)
```python
import requests

def convert_ppt(file_path):
    url = 'http://localhost:5001/api/convert'
    
    with open(file_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(url, files=files)
    
    if response.status_code == 200:
        result = response.json()
        if result['success']:
            return result['data']
        else:
            raise Exception(result['message'])
    else:
        raise Exception(f'HTTP {response.status_code}: {response.text}')
```

### cURL
```bash
# 转换PPT文件
curl -X POST \
  http://localhost:5001/api/convert \
  -H 'Content-Type: multipart/form-data' \
  -F 'file=@presentation.pptx' \
  -o result.json

# 检查服务状态
curl http://localhost:5001/api/health
``` 