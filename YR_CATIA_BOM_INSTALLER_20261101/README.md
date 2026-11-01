# YR CATIA BOM Manager 安装包工程

## 给客户的文件

只需要发送：

`dist\YR_CATIA_BOM_Setup.exe`

客户双击后可以选择“安装 / 更新”或“卸载”。软件按当前 Windows 用户安装，不需要管理员权限。

## 安装行为

- 安装位置：`%LocalAppData%\Programs\YR CATIA BOM Manager`
- 用户输出：`%UserProfile%\Documents\YR CATIA BOM Manager`
- 自动创建桌面和开始菜单快捷方式
- 在 Windows“已安装的应用”中登记卸载入口
- 内置 Python、pywin32、openpyxl、Pillow 等运行组件
- 不修改系统 Python，不设置系统环境变量，不需要联网安装依赖
- 卸载时可选择保留本地文件，或删除 `%UserProfile%\Documents\YR CATIA BOM Manager` 中的 BOM、图片、日志及模板结果（默认保留）

CATIA V5 本体和有效许可证必须由客户电脑原本安装，本安装包不会安装或修改 CATIA。

## 重新构建

双击 `BUILD_INSTALLER.cmd`。构建脚本使用现有 `CATIA_BOM_GUI\.venv`，自动安装本地构建工具，然后依次生成应用程序和单文件安装器。

发布前建议购买代码签名证书，对 `YR_CATIA_BOM_Setup.exe` 和主程序签名，否则部分公司电脑可能显示“未知发布者”或触发安全软件检查。
