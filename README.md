# XLS 导入 Access 小工具

## 功能说明
- 读取 xls/xlsx 工作表
- 自动在 Access 中创建新表并导入数据
- 默认只判断日期时间列，其余列按文本处理;导入失败可试试“全文本存入”
- 表名默认使用工作表名，如冲突自动补时间戳
- 大文件按批量导入，默认每批 1000 行

## 使用前提
- 机器上需要可用的 Microsoft Access Database Engine / ACE OLE DB 驱动
  > **Microsoft.ACE.OLEDB.12.0** https://www.microsoft.com/en-us/download/details.aspx?id=13255
  > **Microsoft.ACE.OLEDB.16.0** https://www.microsoft.com/en-us/download/details.aspx?id=54920
- 如果目标 .mdb 文件不存在，程序会自动创建一个空白 .mdb 后再导入
- 如果使用已有 .accdb 或 .mdb 文件，程序会直接在其中建表导入


## 使用方式
1. 选择 Excel 文件、工作表
2. 如果已有库，点击“打开已有”选择 ；默认在同目录新建`.mdb`，也可以点击“新建 MDB”自定义
3. 确认目标表名、字段类型和批量大小

## 注意事项
- Excel 第一行会作为字段名
- 日期时间列若存在无法转换的值，该行会失败并写入日志
- 如果数据较乱，建议勾选“全文本存入”后再导入
- 自动创建数据库时当前只支持 `.mdb`
- 大文件会先读取预览再分批导入，预览行数不代表总行数
