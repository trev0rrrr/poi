# poi
## 介绍
>apache poi是Apache的开源项目，提供Java对Microsoft office格式档案的读写功能。  

## 结构/定义


poi名称 | 操作对象
---- | ----
HSSF | Excel xls格式
XSSF | Excel ooxml格式
HWPF | Word doc格式
HSLF | ppt
HDGF | visio
HPBF | publisher
HSMF | Outlook

主要使用excel做报表，介绍一下相关类

类名 | desc
---- | ----
HSSFWorkbook    | Excel文档对象
HSSFSheet       | 表单
HSSFRow         | 行
HSSFCell        | 单元格
HSSFFont        | 字体
HSSFDateFormat  | 单元格的日期格式
HSSFHeader      | sheet页眉
HSSFFooter      | sheet页脚
HSSFCellStyle   | 单元格样式


