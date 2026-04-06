// =============================================================================
// CorelDRAW COM API — TextRange 获取帮助类
// 适用版本：CorelDRAW X7 及以上
// 语言：C# 8.0+（.NET Framework 4.7.2 / .NET 6+）
//
// 核心 API（来自官方《CorelDRAW X7 脚本参考手册》第 5 册）：
//   Text.Selection  — 编辑模式下光标当前选中区域 → TextRange
//   Text.Story      — 整个文字流的完整 TextRange（跨帧）
//   Text.IsEditing  — 是否处于文字编辑模式
//   TextRange.Length — 选区字符数（0 = 无选区）
//   Page.FindShapes(Name, Type, Query) — 按类型筛选形状
//
// 使用方式（后期绑定，无需额外 DLL 引用）：
//   dynamic corelApp = Marshal.GetActiveObject("CorelDRAW.Application");
//   var ranges = CorelTextRangeHelper.GetTextRanges(corelApp);
//   foreach (var tr in ranges) tr.Font = "Arial";
//
// 若使用早期绑定（添加 Corel.Interop.VGCore COM 引用），可将所有
// dynamic 替换为对应的强类型接口（如 Application、TextRange 等）。
// =============================================================================

using System;
using System.Collections.Generic;

namespace CorelDrawCOMAPI
{
    /// <summary>
    /// 指定 <see cref="CorelTextRangeHelper"/> 获取 TextRange 的范围来源。
    /// </summary>
    public enum TextRangeScope
    {
        /// <summary>
        /// 自动检测，按以下优先级依次尝试：
        /// 编辑模式选区 → 编辑模式 Story → 选中形状 → 当前页面所有文本。
        /// </summary>
        Auto = 0,

        /// <summary>
        /// 当前文字编辑模式下光标选中的文字区域（<c>Text.Selection</c>）。
        /// 仅在用户双击进入文本框并选中文字时才有内容。
        /// </summary>
        EditingSelection = 1,

        /// <summary>
        /// 当前选区中所有文本形状（<c>cdrTextShape</c>）各自的完整文字（<c>Text.Story</c>）。
        /// </summary>
        SelectedShapes = 2,

        /// <summary>
        /// 当前活动页面（或通过 <c>page</c> 参数指定的页面）中所有文本形状的完整文字。
        /// </summary>
        CurrentPage = 3,

        /// <summary>
        /// 当前文档（或通过 <c>doc</c> 参数指定的文档）所有页面中所有文本形状的完整文字。
        /// </summary>
        CurrentDocument = 4,
    }

    /// <summary>
    /// CorelDRAW COM API — TextRange 获取帮助类（静态工具类）。
    /// <para>
    /// 封装了从四种来源安全获取 <c>TextRange</c> 对象的逻辑：
    /// 编辑模式选区、选中形状、当前页面、整个文档。
    /// 所有公共方法均不会抛出 COM 异常；
    /// 集合方法在找不到文本时返回空列表（不返回 <c>null</c>）；
    /// 单对象方法在找不到目标时返回 <c>null</c>。
    /// </para>
    /// <para>
    /// <b>依赖</b>：运行时需要 CorelDRAW X7 或更高版本已安装并注册。
    /// 使用后期绑定（<c>dynamic</c>），无需在项目中添加 COM 引用即可编译。
    /// </para>
    /// </summary>
    public static class CorelTextRangeHelper
    {
        // cdrShapeType.cdrTextShape（CorelDRAW X7 类型库中的枚举值）
        private const int CdrTextShape = 9;

        // ─────────────────────────────── 主入口 ─────────────────────────────────

        /// <summary>
        /// 根据 <paramref name="scope"/> 返回对应的所有 <c>TextRange</c> 列表。
        /// <para>返回值永远不为 <c>null</c>；没有文本时返回空列表。</para>
        /// </summary>
        /// <param name="app">
        /// CorelDRAW <c>Application</c> COM 对象（<c>dynamic</c>）。
        /// 可通过 <c>Marshal.GetActiveObject("CorelDRAW.Application")</c> 获取。
        /// </param>
        /// <param name="scope">
        /// 获取范围，默认 <see cref="TextRangeScope.Auto"/> 自动检测。
        /// </param>
        /// <returns>
        /// 包含零个或多个 <c>TextRange</c> COM 对象的只读列表。
        /// 每个元素均可直接访问 <c>.Font</c>、<c>.Size</c>、<c>.Bold</c> 等属性。
        /// </returns>
        /// <example>
        /// <code>
        /// dynamic app = Marshal.GetActiveObject("CorelDRAW.Application");
        /// foreach (var tr in CorelTextRangeHelper.GetTextRanges(app))
        /// {
        ///     tr.Font = "微软雅黑";
        ///     tr.Size = 12;
        /// }
        /// </code>
        /// </example>
        public static IReadOnlyList<dynamic> GetTextRanges(
            dynamic app,
            TextRangeScope scope = TextRangeScope.Auto)
        {
            if (app == null) return Array.Empty<dynamic>();

            try
            {
                switch (scope)
                {
                    case TextRangeScope.EditingSelection: return GetEditingSelectionList(app);
                    case TextRangeScope.SelectedShapes:  return GetSelectedShapeStoriesList(app);
                    case TextRangeScope.CurrentPage:     return GetPageStoriesList(app, null);
                    case TextRangeScope.CurrentDocument: return GetDocumentStoriesList(app, null);
                    default:                             return AutoDetect(app);
                }
            }
            catch
            {
                return Array.Empty<dynamic>();
            }
        }

        /// <summary>
        /// 返回单个最相关的 <c>TextRange</c>，没有可用文本时返回 <c>null</c>。
        /// <para>等价于取 <see cref="GetTextRanges"/> 返回列表的第一个元素。</para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="scope">获取范围，默认自动检测。</param>
        /// <returns>第一个可用的 <c>TextRange</c>；或 <c>null</c>。</returns>
        /// <example>
        /// <code>
        /// dynamic tr = CorelTextRangeHelper.TryGetTextRange(app);
        /// if (tr != null) tr.Bold = true;
        /// </code>
        /// </example>
        public static dynamic TryGetTextRange(
            dynamic app,
            TextRangeScope scope = TextRangeScope.Auto)
        {
            var list = GetTextRanges(app, scope);
            return list.Count > 0 ? list[0] : null;
        }

        // ─────────────────────────────── 分场景方法 ─────────────────────────────

        /// <summary>
        /// 获取当前文字编辑模式下光标选中的 <c>TextRange</c>（对应 <c>Text.Selection</c>）。
        /// <para>
        /// 前提：用户双击进入文字框且选中了至少一个字符。
        /// 仅处于编辑模式但未选文字（光标是插入点）时返回 <c>null</c>。
        /// </para>
        /// </summary>
        /// <returns>选中的 TextRange；或 <c>null</c>。</returns>
        public static dynamic GetEditingSelection(dynamic app)
        {
            var list = GetEditingSelectionList(app);
            return list.Count > 0 ? list[0] : null;
        }

        /// <summary>
        /// 获取当前选区中所有文本形状各自的完整文字（<c>Text.Story</c>）。
        /// <para>
        /// 对应 VBA：
        /// <code>For Each s In ActiveSelectionRange : s.Text.Story : Next</code>
        /// </para>
        /// </summary>
        /// <returns>选区内所有文本形状的 Story TextRange 列表（可能为空）。</returns>
        public static IReadOnlyList<dynamic> GetSelectedShapeStories(dynamic app)
            => GetSelectedShapeStoriesList(app);

        /// <summary>
        /// 获取指定页面（默认活动页面）中所有文本形状的完整文字（<c>Text.Story</c>）。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="page">
        /// 目标 Page COM 对象；传 <c>null</c> 则使用 <c>app.ActivePage</c>。
        /// </param>
        /// <returns>页面上所有文本形状的 Story TextRange 列表（可能为空）。</returns>
        public static IReadOnlyList<dynamic> GetPageTextStories(
            dynamic app,
            dynamic page = null)
            => GetPageStoriesList(app, page);

        /// <summary>
        /// 获取指定文档（默认活动文档）所有页面中所有文本形状的完整文字（<c>Text.Story</c>）。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="doc">
        /// 目标 Document COM 对象；传 <c>null</c> 则使用 <c>app.ActiveDocument</c>。
        /// </param>
        /// <returns>文档内所有文本形状的 Story TextRange 列表（可能为空）。</returns>
        public static IReadOnlyList<dynamic> GetDocumentTextStories(
            dynamic app,
            dynamic doc = null)
            => GetDocumentStoriesList(app, doc);

        /// <summary>
        /// 对从指定范围获取的每一个 <c>TextRange</c> 执行 <paramref name="action"/>。
        /// <para>单个 action 抛出的异常会被捕获并跳过，不影响后续处理。</para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="action">对每个 TextRange 执行的操作，参数类型为 <c>dynamic</c>。</param>
        /// <param name="scope">获取范围，默认自动检测。</param>
        /// <example>
        /// <code>
        /// // 将所有选中文本设为 Arial 14pt 粗体
        /// CorelTextRangeHelper.ForEachTextRange(app, tr =>
        /// {
        ///     tr.Font = "Arial";
        ///     tr.Size = 14;
        ///     tr.Bold = true;
        /// }, TextRangeScope.SelectedShapes);
        /// </code>
        /// </example>
        public static void ForEachTextRange(
            dynamic app,
            Action<dynamic> action,
            TextRangeScope scope = TextRangeScope.Auto)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));

            foreach (var tr in GetTextRanges(app, scope))
            {
                try { action(tr); }
                catch { /* 单个 TextRange 操作失败，跳过继续 */ }
            }
        }

        // ─────────────────────────────── 私有实现 ───────────────────────────────

        private static IReadOnlyList<dynamic> AutoDetect(dynamic app)
        {
            // 优先级 1：文字编辑模式且有选中字符
            var editSel = GetEditingSelectionList(app);
            if (editSel.Count > 0)
                return editSel;

            // 优先级 2：文字编辑模式但无选区 → 返回整个 Story（插入点所在文字框）
            dynamic story = GetEditingStory(app);
            if (story != null)
                return new[] { story };

            // 优先级 3：用选择工具选中了文字形状
            var selected = GetSelectedShapeStoriesList(app);
            if (selected.Count > 0)
                return selected;

            // 优先级 4：兜底，返回当前页面所有文本
            return GetPageStoriesList(app, null);
        }

        private static IReadOnlyList<dynamic> GetEditingSelectionList(dynamic app)
        {
            try
            {
                dynamic shape = app.ActiveShape;
                if (shape == null) return Array.Empty<dynamic>();
                if ((int)shape.Type != CdrTextShape) return Array.Empty<dynamic>();

                dynamic text = shape.Text;
                if (!(bool)text.IsEditing) return Array.Empty<dynamic>();

                dynamic sel = text.Selection;
                if (sel == null || (int)sel.Length == 0) return Array.Empty<dynamic>();

                return new[] { sel };
            }
            catch
            {
                return Array.Empty<dynamic>();
            }
        }

        private static dynamic GetEditingStory(dynamic app)
        {
            try
            {
                dynamic shape = app.ActiveShape;
                if (shape == null) return null;
                if ((int)shape.Type != CdrTextShape) return null;

                dynamic text = shape.Text;
                if (!(bool)text.IsEditing) return null;

                return text.Story;
            }
            catch
            {
                return null;
            }
        }

        private static IReadOnlyList<dynamic> GetSelectedShapeStoriesList(dynamic app)
        {
            var result = new List<dynamic>();
            try
            {
                dynamic selRange = app.ActiveSelectionRange;
                if (selRange == null) return result;

                foreach (dynamic shape in selRange)
                {
                    try
                    {
                        if ((int)shape.Type == CdrTextShape)
                            result.Add(shape.Text.Story);
                    }
                    catch { /* 跳过无法访问的形状 */ }
                }
            }
            catch { /* 无选区时忽略 */ }
            return result;
        }

        private static IReadOnlyList<dynamic> GetPageStoriesList(dynamic app, dynamic page)
        {
            var result = new List<dynamic>();
            try
            {
                dynamic targetPage = page ?? app.ActivePage;
                if (targetPage == null) return result;

                // FindShapes(Name, Type, Query)：第 2 个参数 Type = cdrTextShape(9)
                dynamic shapes = targetPage.FindShapes(null, CdrTextShape);
                foreach (dynamic shape in shapes)
                {
                    try { result.Add(shape.Text.Story); }
                    catch { }
                }
            }
            catch { }
            return result;
        }

        private static IReadOnlyList<dynamic> GetDocumentStoriesList(dynamic app, dynamic doc)
        {
            var result = new List<dynamic>();
            try
            {
                dynamic targetDoc = doc ?? app.ActiveDocument;
                if (targetDoc == null) return result;

                foreach (dynamic page in targetDoc.Pages)
                {
                    try
                    {
                        dynamic shapes = page.FindShapes(null, CdrTextShape);
                        foreach (dynamic shape in shapes)
                        {
                            try { result.Add(shape.Text.Story); }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return result;
        }
    }
}
