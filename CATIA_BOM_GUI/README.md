# CATIA BOM Manager GUI

这是基于现有 CATIA BOM 导出、截图和基础数据管制表转换脚本的 Windows GUI 第一版。

## 启动

首次使用请双击 `SETUP_GUI_RUNTIME.cmd`。它会使用系统安装的标准 CPython 3.10+ 创建本软件专用的 `.venv` 环境，并安装 `pywin32`、`openpyxl`、`Pillow`。不要用 `C:\Espressif\tools\python` 这类 ESP-IDF 专用 Python 运行本软件。

安装完成后，双击 `START_GUI.cmd` 启动。

也可在 Windows 终端进入本目录后手动执行：

```powershell
py -3.10 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\catia_bom_gui.py
```

启动器只含英文 ASCII 字符，避免不同 Windows 编码导致的批处理闪退。

## 已实现

- 导出当前 CATIA 活动 `.CATProduct`；
- 选择多个根 `.CATProduct` 后串行批量导出；
- 在“导出 BOM”页选择客户模板后，可一键完成 BOM 导出和模板转换；
- 自动转换结果可指定路径，也可留空后按本次 BOM 自动命名；批量模式会在全部任务导出成功后，把本批次 BOM 合并到同一份模板；
- 截图、材质、质量、名义厚度、包装包络和 Excel BOM；
- 导出页可以选择“长城数据”：总成、分总成和零件的名称均读取各自产品引用属性的 `V_description`；零件的版本、材质、重量与三维尺寸也从属性读取。仅当材料分类为钢板且材料规格是有效数值时，将材料规格作为待图纸确认的名义厚度。缺失值留空，不推断几何板厚；长城属性规则集中在 `gwm_properties.py`。
- 截图时隐藏总成、子总成及零件内的 PMI/FTA 标注，导出结束后仍保持隐藏；`CAD_Deliverables`、XY/YZ/ZX 基准面和左侧规格树在导出后恢复原有显示状态；
- 大总成截图会在首次产品树遍历时缓存叶子实例，后续按相邻截图的可见差量批量切换，避免每张图重复扫描完整产品树；
- 在“转换模板”页一次添加一份或多份图片清单 BOM，并合并写入**同一份**客户基础数据管制表模板；
- 合并时会为每个根总成重新编号，因此不同 BOM 的层级、图片和重量汇总不会相互串联；
- 实时日志、进度条、输出目录入口；
- 安全取消：当前 CATIA 测量或截图完成后停止，不保存客户源文件。

## 使用约束

- 单件导出前，必须在 CATIA 中打开并激活 `.CATProduct`。
- 批量导出只应添加客户指定的**根总成**，不要把同一总成内的子总成重复加入。
- 合并转换时，只添加需要交付给同一客户模板的根总成 BOM；不要同时加入同一 BOM 的重复副本。
- CATIA 与本程序应以相同 Windows 用户和相同权限运行。
- 本版本是开发交付版；正式商用前应增加安装包、许可证、代码签名、回归测试与客户模板配置管理。

## 源码位置

- GUI 和 BOM 导出内核统一位于本目录：`catia_bom_gui.py`、`catia_bom_export_v2.py`。
- 仓库根目录的 `catia_bom_export_v2.py` 仅为兼容旧命令行启动方式的入口；后续功能修改应只改本目录内的文件。
