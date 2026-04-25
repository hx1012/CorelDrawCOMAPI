# CorelDRAW COM API — TextRange 获取工具类（C#）

> 内容来源：《CorelDRAW X7 脚本参考手册》第 5 册  
> 文件：[`TextRangeHelper.cs`](TextRangeHelper.cs)

---

## 目录

1. [背景与需求](#一背景与需求)
2. [环境要求](#二环境要求)
3. [快速开始](#三快速开始)
4. [API 速查](#四api-速查)
   - 4.1 [TextRangeScope 枚举](#41-textrangescope-枚举)
   - 4.2 [CorelTextRangeHelper 方法一览](#42-coreltextrangehelper-方法一览)
5. [分场景详解与示例](#五分场景详解与示例)
   - 5.1 [场景一：编辑模式中的选中文字](#51-场景一编辑模式中的选中文字)
   - 5.2 [场景二：当前选中的文字形状](#52-场景二当前选中的文字形状)
   - 5.3 [场景三：当前页面所有文本](#53-场景三当前页面所有文本)
   - 5.4 [场景四：整个文档所有文本](#54-场景四整个文档所有文本)
   - 5.5 [自动检测（推荐）](#55-自动检测推荐)
6. [ForEachTextRange 批量操作](#六foreachtextrange-批量操作)
7. [对象模型说明](#七对象模型说明)
8. [自动检测优先级](#八自动检测优先级)
9. [返回值始终安全](#九返回值始终安全)
10. [早期绑定升级指南](#十早期绑定升级指南)
11. [常见问题](#十一常见问题)
12. [API 参考](#十二api-参考)

---

## 一、背景与需求

在 CorelDRAW COM API 中，`TextRange` 是操作文字内容（字体、字号、颜色、对齐等）的核心对象。  
但 **`TextRange` 的来源可以有四种**，每种的获取路径不同：

| 来源 | 用户操作 | VBA 路径 |
|------|----------|---------|
| 编辑模式选区 | 双击文字框并拖选文字 | `ActiveShape.Text.Selection` |
| 选中文字形状 | 用选择工具点选/框选文字对象 | `ActiveSelectionRange` → `.Text.Story` |
| 当前页面所有文字 | 无特定选区 | `ActivePage.FindShapes(Type:=9)` → `.Text.Story` |
| 文档所有文字 | 无特定选区 | 遍历 `ActiveDocument.Pages` |

`CorelTextRangeHelper` 将上述四种情况统一封装，提供**一个统一入口 `GetTextRanges()`**，  
调用时无需关心当前是哪种状态，也无需手写 `null` 检查和 `try/catch`。

---

## 二、环境要求

| 项目 | 要求 |
|------|------|
| CorelDRAW | X7 或以上版本（X8、2017–2021 均可） |
| .NET 框架 | .NET Framework 4.7.2 / .NET 6 / .NET 8 |
| C# 版本 | C# 8.0+（使用了 `null` 合并运算符） |
| COM 引用 | **不需要**（使用 `dynamic` 后期绑定） |

> **早期绑定可选**：若项目中已添加 `Corel.Interop.VGCore` COM 引用，可将所有 `dynamic`  
> 替换为强类型接口，获得 IntelliSense 支持和编译期类型检查。见[第十节](#十早期绑定升级指南)。

---

## 三、快速开始

### 1. 将 `TextRangeHelper.cs` 加入项目

直接将 [`TextRangeHelper.cs`](TextRangeHelper.cs) 复制到你的 C# 项目中即可。  
无需 NuGet 包，无需 COM 引用。

### 2. 引入命名空间

```csharp
using System.Runtime.InteropServices;
using CorelDrawCOMAPI;
```

### 3. 获取 CorelDRAW Application 对象

```csharp
// 连接到已运行的 CorelDRAW 实例
dynamic app = Marshal.GetActiveObject("CorelDRAW.Application");
```

### 4. 调用帮助类

```csharp
// 自动检测来源，批量修改字体
foreach (dynamic tr in CorelTextRangeHelper.GetTextRanges(app))
{
    tr.Font = "微软雅黑";
    tr.Size = 12;
}

// 或使用 ForEachTextRange（更简洁，异常自动跳过）
CorelTextRangeHelper.ForEachTextRange(app, tr =>
{
    tr.Font = "Arial";
    tr.Bold = true;
});
```

---

## 四、API 速查

### 4.1 `TextRangeScope` 枚举

```csharp
public enum TextRangeScope
{
    Auto             = 0,  // 自动检测（默认）
    EditingSelection = 1,  // 编辑模式中被选中的文字区域
    SelectedShapes   = 2,  // 选区中所有文字形状的 Story
    CurrentPage      = 3,  // 当前页面所有文字形状的 Story
    CurrentDocument  = 4,  // 整个文档所有文字形状的 Story
}
```

### 4.2 `CorelTextRangeHelper` 方法一览

| 方法 | 返回值 | 说明 |
|------|--------|------|
| `GetTextRanges(app, scope)` | `IReadOnlyList<dynamic>` | 主入口，永不返回 null，找不到时返回空列表 |
| `TryGetTextRange(app, scope)` | `dynamic`（可为 null） | 取列表第一个元素；空列表时返回 null |
| `GetEditingSelection(app)` | `dynamic`（可为 null） | 编辑模式下的光标选区 |
| `GetSelectedShapeStories(app)` | `IReadOnlyList<dynamic>` | 选中形状各自的 Story |
| `GetPageTextStories(app, page)` | `IReadOnlyList<dynamic>` | 指定页面（null=当前页）的所有 Story |
| `GetDocumentTextStories(app, doc)` | `IReadOnlyList<dynamic>` | 指定文档（null=当前文档）的所有 Story |
| `ForEachTextRange(app, action, scope)` | `void` | 对每个 TextRange 执行委托，异常自动跳过 |

---

## 五、分场景详解与示例

### 5.1 场景一：编辑模式中的选中文字

**触发条件**：用户双击文字框，进入编辑模式并拖选了一段文字。  
**API 路径**：`ActiveShape.Text.Selection`（仅在 `Text.IsEditing = true` 且 `Selection.Length > 0` 时有效）

```csharp
// 方式 A：明确指定 EditingSelection
var ranges = CorelTextRangeHelper.GetTextRanges(app, TextRangeScope.EditingSelection);
if (ranges.Count > 0)
{
    ranges[0].Bold = true;
    ranges[0].Fill.ApplyUniformFill(CreateRGBColor(255, 0, 0));  // 变红
}

// 方式 B：获取单个 TextRange
dynamic tr = CorelTextRangeHelper.GetEditingSelection(app);
if (tr != null)
    tr.Size = 18;
```

> **注意**：如果用户处于编辑模式但没有选中文字（仅光标插入点），  
> `EditingSelection` 返回空列表。自动检测（`Auto`）此时会回退到整个 Story。

---

### 5.2 场景二：当前选中的文字形状

**触发条件**：用户用选择工具（而非编辑模式）选中了一个或多个文字对象。  
**API 路径**：遍历 `ActiveSelectionRange`，对 `cdrTextShape` 类型的形状取 `.Text.Story`

```csharp
var stories = CorelTextRangeHelper.GetSelectedShapeStories(app);
Console.WriteLine($"选中了 {stories.Count} 个文字形状");

foreach (dynamic tr in stories)
{
    tr.Font      = "Arial";
    tr.Size      = 14;
    tr.Alignment = 1;  // cdrLeftAlignment = 1
}
```

---

### 5.3 场景三：当前页面所有文本

```csharp
// 当前活动页面
var stories = CorelTextRangeHelper.GetPageTextStories(app);

// 指定页面（如第 2 页）
dynamic page2 = app.ActiveDocument.Pages[2];
var page2Stories = CorelTextRangeHelper.GetPageTextStories(app, page2);

foreach (dynamic tr in stories)
    tr.Font = "黑体";
```

---

### 5.4 场景四：整个文档所有文本

```csharp
// 当前文档
var allStories = CorelTextRangeHelper.GetDocumentTextStories(app);
Console.WriteLine($"文档共有 {allStories.Count} 个文字对象");

// 指定文档
dynamic doc = app.Documents[1];
var docStories = CorelTextRangeHelper.GetDocumentTextStories(app, doc);

foreach (dynamic tr in docStories)
{
    tr.Font = "宋体";
    tr.Size = 10;
}
```

---

### 5.5 自动检测（推荐）

**不确定用户当前状态时，直接用 `Auto`（默认值），让帮助类自动判断：**

```csharp
// 自动检测：内部按优先级逐级尝试，始终能返回有意义的结果
var ranges = CorelTextRangeHelper.GetTextRanges(app);  // scope = Auto

foreach (dynamic tr in ranges)
{
    tr.Font = "微软雅黑";
    tr.Size = 12;
}
```

等价的 `TryGetTextRange` 用法（只需操作一个 TextRange 时）：

```csharp
dynamic tr = CorelTextRangeHelper.TryGetTextRange(app);
if (tr != null)
{
    tr.Bold      = true;
    tr.Underline = 1;  // cdrSingleThinFontLine = 1
}
```

---

## 六、`ForEachTextRange` 批量操作

`ForEachTextRange` 是最简洁的批量操作入口，内置异常保护：

```csharp
// 将所有选中形状的文字统一设为 Arial 粗体
CorelTextRangeHelper.ForEachTextRange(app, tr =>
{
    tr.Font = "Arial";
    tr.Bold = true;
}, TextRangeScope.SelectedShapes);

// 将当前文档所有文字颜色设为黑色
CorelTextRangeHelper.ForEachTextRange(app, tr =>
{
    tr.Fill.ApplyUniformFill(
        app.CreateRGBColor(0, 0, 0));
}, TextRangeScope.CurrentDocument);
```

---

## 七、对象模型说明

```
Application
└── ActiveShape              → Shape (cdrTextShape)
    └── .Text                → Text 对象
        ├── .Story           → TextRange（整个文字流，最常用）
        ├── .Selection       → TextRange（编辑模式光标选区）
        ├── .Range(s, e)     → TextRange（按字符索引取子范围）
        ├── .IsEditing       → bool（是否处于编辑模式）
        └── .Frames          → TextFrames（段落文字帧集合）
            └── [n].Range    → TextRange（某一帧内的文字）

Application
└── ActiveSelectionRange     → ShapeRange（当前选区）
    └── [cdrTextShape 类型]  → Shape → .Text.Story → TextRange

Page / Document
└── FindShapes(null, 9)      → ShapeRange（页面/文档中所有文字形状）
    └── .Text.Story          → TextRange
```

**TextRange 常用属性**（获取后可直接操作）：

| 属性 | 类型 | 说明 |
|------|------|------|
| `Font` | `string` | 字体名称 |
| `Size` | `float` | 字号（磅） |
| `Bold` | `bool` | 粗体 |
| `Italic` | `bool` | 斜体 |
| `Underline` | `int`（cdrFontLine） | 下划线样式 |
| `Strikethru` | `int`（cdrFontLine） | 删除线样式 |
| `Alignment` | `int`（cdrAlignment） | 对齐方式 |
| `Fill` | `Fill` 对象 | 文字颜色/填充 |
| `CharSpacing` | `float` | 字符间距 |
| `Length` | `int` | 字符数 |
| `Text` | `string` | 文字内容（可读写） |
| `Words` | `TextWords` | 词集合 |
| `Lines` | `TextLines` | 行集合 |
| `Paragraphs` | `TextParagraphs` | 段落集合 |
| `SetRange(s, e)` | 方法 | 重新定位范围 |
| `Range(s, e)` | 方法 | 返回子 TextRange |

---

## 八、自动检测优先级

`TextRangeScope.Auto`（默认）按以下顺序检测，第一个命中即返回：

```
1. 文字编辑模式 + 有选中字符
   → 返回 Text.Selection（光标选区）

2. 文字编辑模式 + 无选中字符（仅插入点）
   → 返回 Text.Story（当前文字框的完整文字）

3. 未在编辑模式，但选区中有文字形状
   → 返回选区内所有文字形状的 Story 列表

4. 以上均不满足（无选区或选区中无文字）
   → 返回当前页面中所有文字形状的 Story 列表
```

---

## 九、返回值始终安全

所有方法内部均包含完整的 `try/catch`，不会向外抛出 COM 异常。

| 方法类型 | 无可用文本时 | 说明 |
|---------|------------|------|
| 集合方法 `GetTextRanges`、`Get*Stories` | 返回空 `IReadOnlyList`（非 null） | 可安全 `foreach`，不需要 null 检查 |
| 单对象方法 `TryGetTextRange`、`GetEditingSelection` | 返回 `null` | 调用前需检查 `!= null` |
| 遍历方法 `ForEachTextRange` | 不执行 action | 无需任何前置检查 |

---

## 十、早期绑定升级指南

如需 IntelliSense 和编译期类型安全，可在项目中添加 CorelDRAW COM 引用：

1. 项目 → 添加引用 → COM → 搜索 **CorelDRAW Application**（Corel.Interop.VGCore）
2. 将文件中的 `dynamic app` 替换为：
   ```csharp
   using VGCore = Corel.Interop.VGCore;
   VGCore.Application app = ...;
   ```
3. 将返回值 `dynamic` 替换为 `VGCore.TextRange`：
   ```csharp
   // 原来
   dynamic tr = CorelTextRangeHelper.TryGetTextRange(app);
   // 强类型
   VGCore.TextRange tr = (VGCore.TextRange)CorelTextRangeHelper.TryGetTextRange(app);
   ```

> 在 `TextRangeHelper.cs` 内部的私有方法中，`dynamic` 仍可保留，以避免大量转型代码。

---

## 十一、常见问题

**Q：`GetTextRanges` 返回空列表，但页面上明明有文字？**  
A：检查 CorelDRAW 是否处于运行状态，且 `Marshal.GetActiveObject` 已正确连接。  
也可能是文字形状被锁定在不可访问的图层上。

**Q：`GetEditingSelection` 始终返回 null？**  
A：需要用户**双击**进入文字框编辑模式，并**拖选了字符**。仅单击进入（光标插入点）时 `Length = 0`，此时 `Auto` 会自动退回返回整个 Story。

**Q：修改 `tr.Font` 后 CorelDRAW 没有变化？**  
A：文字修改需要文档处于可编辑状态。若文档受保护或图层被锁定，COM 赋值会静默失败。

**Q：能否跨多个文字对象返回一个"合并的" TextRange？**  
A：不能。CorelDRAW 的 TextRange 只能属于同一个文字形状（Text 对象），  
无法跨 Shape 合并。`GetTextRanges` 返回的是每个形状各自的 Story 列表，需分别操作。

**Q：如何只操作 Story 中的第 3～10 个字符？**  
```csharp
dynamic tr = CorelTextRangeHelper.TryGetTextRange(app);
if (tr != null)
{
    dynamic sub = tr.Range(3, 10);  // TextRange.Range(start, end)
    sub.Bold = true;
    // 或用 SetRange 就地修改范围
    tr.SetRange(3, 10);
    tr.Underline = 1;
}
```

---

## 十二、API 参考

| 手册来源 | 页码 |
|---------|------|
| `Text.Selection`（属性） | 第 5 册第 11367 页 |
| `Text.Story`（属性） | 第 5 册第 11385 页 |
| `Text.IsEditing`（属性） | 第 5 册第 11168 页 |
| `TextRange.Length`（属性） | 第 5 册第 14464 页 |
| `TextRange.Range()`（方法） | 第 5 册第 14798 页 |
| `TextRange.SetRange()`（方法） | 第 5 册第 15043 页 |
| `Page.FindShapes()`（方法） | 第 3 册 |
| `cdrShapeType`（枚举） | 第 1 册（`cdrTextShape = 9`） |

---

## 参考资料

- 本仓库中的《CorelDraw脚本参考手册X7》PDF（共 8 册）
- [`字体样式设置指南.md`](字体样式设置指南.md)：TextRange 字体属性完整说明
- [`Shape对象完整指南.md`](Shape对象完整指南.md)：SelectionInfo、FindShapes 等选区 API
