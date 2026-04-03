# CorelDRAW X7 COM API 脚本参考

本仓库整理了通过阅读官方《CorelDRAW X7 脚本参考手册》（PDF / CHM）提炼的 VBA 宏使用指南与示例代码，涵盖文字样式设置、表格创建与格式化等常用场景，方便开发者快速查阅正确的 COM API 用法。

---

## 仓库文件说明

| 文件 / 目录 | 说明 |
|-------------|------|
| `字体样式设置指南.md` | 字体名称、字号、粗/斜/下划线、大小写、上下标、颜色、间距等完整 API 速查 |
| `创建表格指南.md` | 表格创建、单元格读写、合并、填充、边框、行列管理等 API 速查及示例 |
| `CreateTable示例.vba` | 可运行的 VBA 宏（产品信息表 + 2009 年 1 月日历两个示例） |
| `CorelDRAW X7 脚本手册.chm` | CorelDRAW X7 官方脚本参考手册（CHM 格式） |
| `CorelDraw脚本参考手册X7_0*.pdf` | CorelDRAW X7 脚本参考手册分册（PDF 格式，共 8 册） |

---

## 运行环境

- **软件**：CorelDRAW X7 或更高版本（X8、2017–2021）
- **语言**：VBA（Visual Basic for Applications）
- **权限**：无需安装额外插件，使用 CorelDRAW 内置宏引擎即可

---

## 快速开始

1. 打开 CorelDRAW。
2. 点击菜单 **工具（Tools）→ 宏（Macros）→ 宏编辑器（Macro Editor）**，或按 **Alt+F11**。
3. 在左侧"工程"窗格中插入新模块，将 `.vba` 文件中的代码粘贴进去。
4. 将光标置于任意 `Sub` 过程内，按 **F5** 运行。
5. 切换回 CorelDRAW 主窗口查看结果。

---

## 内容导览

### 字体样式设置 → [`字体样式设置指南.md`](字体样式设置指南.md)

涵盖以下 `TextRange` 属性与方法：

- **字体 / 字号**：`Font`、`Size`
- **粗体 / 斜体**：`Bold`、`Italic`
- **字体样式枚举**：`Style`（`cdrFontStyle`，含 18 个常量）
- **下划线 / 删除线**：`Underline`、`Strikethru`（`cdrFontLine` 枚举）
- **大小写**：`Case`（`cdrFontCase`）、`ChangeCase`
- **上标 / 下标**：`Position`（`cdrFontPosition`）
- **文字颜色**：`Fill.ApplyUniformFill`
- **对齐方式**：`Alignment`
- **字符 / 词 / 行间距**：`CharSpacing`、`WordSpacing`、`SetLineSpacing`
- **字偶间距 / 旋转 / 偏移**：`RangeKerning`、`CharAngle`、`HorizShift`、`VertShift`

### 创建表格 → [`创建表格指南.md`](创建表格指南.md)

涵盖以下核心 API：

- **创建表格**：`CreateCustomShape("Table", ...)`（⚠️ 不存在 `CreateTable()` 方法）
- **访问 TableShape**：`Shape.Custom`
- **单元格访问**：`TableShape.Cell(Column, Row)`（列在前，行在后）
- **单元格文字**：`TableCell.TextShape.Text.Story`
- **填充 / 合并 / 边框 / 行列管理**

---

## 参考资料

- 本仓库中的《CorelDraw脚本参考手册X7》PDF（共 8 册）
- 本仓库中的《CorelDRAW X7 脚本手册》（CHM）
- CorelDRAW 官方开发者社区：https://community.coreldraw.com/sdk/

