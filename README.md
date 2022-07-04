# SpreadJS_ExportExcelCheckbox
在纯前端在线表格中实现导出 Excel 显示复选框而不是布尔值 功能
# SpreadJS_ExportExcelCheckbox

#### 介绍
在纯前端在线表格中实现导出 Excel 显示复选框而不是布尔值功能

### SpreadJS 示例，导出 Excel 显示复选框而不是布尔值
该示例包括使用 SpreadJS API 的演示脚本，可用于实现导出 Excel 显示复选框而不是布尔值
有关 SpreadJS API 的更多信息，请参阅[SpreadJS API指南]( https://demo.grapecity.com.cn/spreadjs/help/api/) 和[帮助手册]( https://help.grapecity.com.cn/pages/viewpage.action?pageId=5963808)。



### 运行步骤
1、在开始之前，请确保您已满足以下先决条件：
要运行 SpreadJS，浏览器必须支持 HTML5，客户端导入和导出 Excel 需要 IE10及以上。
请先了解 [SpreadJS 的产品使用环境]( https://www.grapecity.com.cn/developer/spreadjs/selection-guide/product-use-environment)，并申请临时部署授权激活
安装并更新NodeJS和NPM
2、克隆或下载此代码库
3、初始化控件，并运行示例脚本
#### 控件初始化
首先，创建一个新页面，并在页面上输入以下代码：
```
<!DOCTYPE html>
    <html>
    <head>
        <title>SpreadJS HTML Test Page</title>
```
2、在页面中添加对 SpreadJS 的引用。代码如下。需要注意的是，SpreadJS 提供压缩过
```
//（minified）的 JavaScript 文件和和用于调试的文件：
<script src="[Your_Scripts_Path]/gc.spread.sheets.all.xxxx.min.js" type="text/javascript"></script>
```
3、添加 CSS 文件以改变Spread.JS 的外观。默认的CSS文件名为： 
gc.spread.sheets.xxxx.css，里面包含了所有的默认样式。该 CSS 文件将会影响滚动条，筛选框及其子元素，单元格和下方标签栏的样式。引入 CSS 的代码如下：
```
//<link href="[Your_CSS_Path]/gc.spread.sheets.xxxx.css" rel="stylesheet" type="text/css"/>
```
4、添加产品授权，代码为（本地测试可以不添加）：
```
GC.Spread.Sheets.LicenseKey = "xxx";
```
5. 添加控件初始化代码。本例会在一个 id 为 “ss” 的 DOM 元素上初始化 SpreadJS：
```
<script type="text/javascript">
// Add your license
// If run this in local for testing, remove or comment below code
 GC.Spread.Sheets.LicenseKey = "xxx";

// Add your code
 window.onload = function(){
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:3});
var sheet = spread.getActiveSheet();
 }
</script>
</head>
<body>
```
6、 创建一个 id 为 “ss” 的元素，SpreadJS 将在该 DOM 中初始化：
```
<div id="ss" style="height: 500px; width: 800px"></div>
</body>
</html>
```
#### 示例代码
```
HTML：
<p>CheckBox导出Excel为复选框而不是布尔值</p>
<button id="export">导出Excel</button>
<div id="ss"></div>

CSS：
#ss{height:600px;width:100%}
p{
    color: #336699;
    text-align: center;
    margin: 12px 0;
}

button{
    margin: 6px 0;
}

JavaScript：
GC.Spread.Common.CultureManager.culture('zh-cn');
window.onload = () => {
    let spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    let sheet = spread.getActiveSheet();
    sheet.setColumnWidth(2,100);
    let c1 = new GC.Spread.Sheets.CellTypes.CheckBox();
    sheet.setCellType(3,2,c1);
    sheet.setValue(3,2,true);
    sheet.setValue(3,3,'是');
    let c2 = new GC.Spread.Sheets.CellTypes.CheckBox();
    sheet.setCellType(4,2,c2);
    sheet.setValue(4,2,false);
    sheet.setValue(4,3,'否');
    let c3 = new GC.Spread.Sheets.CellTypes.CheckBox();
    c3.isThreeState(true);
    c3.textTrue('check');
    c3.textFalse('uncheck');
    c3.textIndeterminate('unknow');
    sheet.setCellType(5,2,c3);
    sheet.setValue(5,2,true);
    
    const toExcel = (spread) => {
            //导出时检查单元格类型
			let spread_copy = new GC.Spread.Sheets.Workbook();
			spread_copy.fromJSON(spread.toJSON())
			let sheet = spread_copy.getActiveSheet()
			let row = sheet.getRowCount()
			let col = sheet.getColumnCount()
			for(let i = 0;i < row;i++){
				for(let j = 0; j < col; j++){
					if(sheet.getCellType(i,j) instanceof GC.Spread.Sheets.CellTypes.CheckBox){
						let text = ''
						let checkbox = sheet.getCellType(i,j)
						if(sheet.getCellType(i,j).isThreeState()){
						;	text = sheet.getValue(i,j) ?  checkbox.textTrue() : (
								sheet.getValue(i,j) == null ? checkbox.textIndeterminate() 
								:
								checkbox.textFalse()
							)
						}
						sheet.setCellType(i,j,null)
						sheet.getValue(i,j) ? sheet.setValue(i,j,'☑ ' + text)  : sheet.setValue(i,j,'☐ ' + text)	
					}
				}
			}
			return spread_copy.toJSON()
		}
		
		const toExcelFile = (spread) => {
		    //导出按钮触发动作
			let excelIO = new GC.Spread.Excel.IO()
			let json = toExcel(spread)
			excelIO.save(json, (blob) => {
			    console.log("hahha")
				saveAs(blob,'checkbox.xlsx')
			},(e) => {
				console.log(e)
			})
		}
    
    $("#export").click(() => {
        toExcelFile(spread)
    })   
}
```
#### 关于 SpreadJS
[SpreadJS]( https://www.grapecity.com.cn/developer/spreadjs) 是一款基于 HTML5 的纯前端表格控件，兼容 450 多种 Excel 公式，具备“高性能、跨平台、与 Excel 高度兼容”的产品特性。使用 SpreadJS，可直接在 Angular、 React、 Vue 等前端框架中实现高效的模板设计、在线编辑和数据绑定等功能，为最终用户提供高度类似 Excel 的使用体验。


