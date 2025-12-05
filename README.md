# 文件导出工具包

一个简单易用的前端文件导出工具包，支持将HTML内容导出为Excel、PDF和Word文档格式。

[GitHub源码仓库地址](https://github.com/xhy12345/html-to-office)

如果觉得好用，欢迎给个Star ⭐️ 支持一下！

## 功能特性

- ✅ **Excel导出** - 将HTML表格导出为.xlsx格式
- ✅ **PDF导出** - 将HTML内容导出为PDF文档（支持多页，兼容html动画绘制）
- ✅ **Word导出** - 将HTML内容导出为.docx格式
- ✅ **多页支持** - PDF导出支持自动分页
- ✅ **数学公式处理** - 自动处理数学公式显示

## 安装

```bash
npm install html-to-office
```

## 快速开始

### 1. 导入工具包

```javascript
import { exportToExcel, htmlToPdf, htmlToDocx } from 'html-to-office';
```

### 2. Excel导出

#### 导出单个表格
```javascript
// 导出id为"table1"的表格
exportToExcel('table1', '用户信息表');
```

### 3. PDF导出

```javascript
// 将id为"content"的HTML元素导出为PDF
htmlToPdf('content', '报告文档');
```

#### 特性说明
- 支持自动分页，长内容会自动分割到多页
- 每页上下预留30px边距
- 自动处理数学公式显示
- 支持高DPI设备，导出清晰图片

### 4. Word文档导出

```javascript
// 将id为"article"的HTML元素导出为Word文档
htmlToDocx('article', '文章文档');
```

## API文档

### exportToExcel(element, name)

将HTML表格导出为Excel文件。

**参数：**
- `element` (string | Array) - 表格元素的ID，或包含多个表格配置的对象数组
- `name` (string, 可选) - 导出的文件名，不包含扩展名。默认为当前时间戳

**返回值：** 导出的Excel文件数据

### htmlToPdf(element, name)

将HTML内容导出为PDF文档。

**参数：**
- `element` (string) - 要导出的HTML元素ID
- `name` (string, 可选) - 导出的文件名，不包含扩展名。默认为当前时间戳

**特性：**
- 自动处理长内容分页
- 每页上下边距30px
- 自动隐藏数学公式辅助元素
- 支持跨域图片导出

### htmlToDocx(element, name)

将HTML内容导出为Word文档。

**参数：**
- `element` (string) - 要导出的HTML元素ID
- `name` (string, 可选) - 导出的文件名，不包含扩展名。默认为当前时间戳

## 技术细节

### 依赖包

- `xlsx` - Excel文件处理
- `html-docx-js-typescript` - HTML转Word
- `file-saver` - 文件保存
- `html2canvas` - HTML转图片
- `jspdf` - PDF生成

### 浏览器兼容性

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## 注意事项

1. **跨域图片**：PDF导出时，如果包含跨域图片，需要确保图片服务器支持CORS
2. **数学公式**：自动处理MathJax生成的数学公式，隐藏辅助元素
3. **分页处理**：长内容会自动分页，但复杂的CSS布局可能需要额外调整
4. **文件大小**：大量数据导出时，建议使用异步操作避免阻塞UI

## 更新日志

### v1.0.0
- ✨ 初始版本发布
- ✅ 支持Excel、PDF、Word导出
- ✅ 支持多表格同时导出
- ✅ 支持多页PDF导出
- ✅ 添加数学公式处理

## 许可证

MIT License