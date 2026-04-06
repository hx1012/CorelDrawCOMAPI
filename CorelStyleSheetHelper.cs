// =============================================================================
// CorelDRAW COM API — StyleSheet / Style / StyleSet 操作帮助类
// 适用版本：CorelDRAW X7 及以上
// 语言：C# 8.0+（.NET Framework 4.7.2 / .NET 6+）
//
// 核心 API（来自官方《CorelDRAW X7 脚本参考手册》第 5 册）：
//   Document.StyleSheet                — 文档样式表入口
//   StyleSheet.CreateStyle(Cat, ...)   — 创建单类别样式
//   StyleSheet.CreateStyleSet(...)     — 创建空样式集容器
//   StyleSheet.CreateStyleFromShape(Shape, Cat, ...) — 从形状捕获单类别样式
//   StyleSheet.CreateStyleSetFromShape(Shape, ...)   — 从形状捕获多类别样式集
//   Style.CreateStyle(Category)        — 在样式集内创建子样式（激活对应属性）
//   Style.Fill / Style.Outline         — 子样式属性（仅创建子样式后才非 null）
//   Style.Delete / Style.Rename        — 删除/重命名
//   StyleSheet.FindStyle(Name)         — 按名称查找
//   Shape.ApplyStyle(Name)             — 应用到形状
//
// ⚠️ 重要说明（ss.Fill 为 null 问题）：
//   CreateStyleSet 创建的是空容器，其 .Fill、.Outline 等属性均为 null。
//   必须先在容器内调用 Style.CreateStyle("fill") / Style.CreateStyle("outline")
//   创建对应的子样式，才能访问对应属性。
//   推荐使用 CreateStyleSetViaShape 方法，通过临时形状间接创建，更为可靠。
//
// 使用方式（后期绑定，无需额外 DLL 引用）：
//   dynamic app = Marshal.GetActiveObject("CorelDRAW.Application");
//   CorelStyleSheetHelper.CreateFillStyle(app, "品牌蓝", 31, 73, 125);
//   CorelStyleSheetHelper.ApplyStyle(app.ActiveShape, "品牌蓝");
// =============================================================================

using System;
using System.Collections.Generic;

namespace CorelDrawCOMAPI
{
    /// <summary>
    /// CorelDRAW COM API — StyleSheet / Style / StyleSet 操作帮助类（静态工具类）。
    /// <para>
    /// 封装了创建、查找、应用和删除 CorelDRAW 命名样式与样式集的全套操作。
    /// 正确处理了 <c>CreateStyleSet</c> 返回对象的 <c>.Fill</c>/<c>.Outline</c> 为 <c>null</c> 的问题。
    /// </para>
    /// <para>
    /// <b>⚠️ 注意</b>：<c>StyleSheet.CreateStyleSet()</c> 创建的是空容器，其
    /// <c>Style.Fill</c>、<c>Style.Outline</c> 等属性均为 <c>null</c>，
    /// 必须先调用 <c>Style.CreateStyle("fill")</c> 等方法创建子样式才可访问。
    /// 本类的 <see cref="CreateStyleSetViaShape"/> 方法通过临时形状绕过此问题。
    /// </para>
    /// <para>
    /// <b>依赖</b>：运行时需要 CorelDRAW X7 或更高版本已安装并注册。
    /// 使用后期绑定（<c>dynamic</c>），无需在项目中添加 COM 引用。
    /// </para>
    /// </summary>
    public static class CorelStyleSheetHelper
    {
        // Category 字符串常量（CorelDRAW X7 样式类别名称，大小写不敏感）
        /// <summary>填充类别名称（<c>"fill"</c>）。</summary>
        public const string CategoryFill      = "fill";
        /// <summary>轮廓类别名称（<c>"outline"</c>）。</summary>
        public const string CategoryOutline   = "outline";
        /// <summary>字符类别名称（<c>"character"</c>）。</summary>
        public const string CategoryCharacter = "character";
        /// <summary>段落类别名称（<c>"paragraph"</c>）。</summary>
        public const string CategoryParagraph = "paragraph";
        /// <summary>图文框类别名称（<c>"frame"</c>）。</summary>
        public const string CategoryFrame     = "frame";

        // ─────────────────────── 样式（单类别）创建 ──────────────────────────

        /// <summary>
        /// 创建一个 RGB 纯色填充样式。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.CreateStyle("fill", "", name, True)</code>
        /// 后设置 <c>Style.Fill.Type</c> 和 <c>Style.Fill.PrimaryColor</c>。
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象（<c>dynamic</c>）。</param>
        /// <param name="name">样式名称。</param>
        /// <param name="r">填充颜色 R（0–255）。</param>
        /// <param name="g">填充颜色 G（0–255）。</param>
        /// <param name="b">填充颜色 B（0–255）。</param>
        /// <param name="replaceExisting">同名样式存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>创建的 <c>Style</c> COM 对象；失败时返回 <c>null</c>。</returns>
        /// <example>
        /// <code>
        /// CorelStyleSheetHelper.CreateFillStyle(app, "品牌蓝", 31, 73, 125);
        /// </code>
        /// </example>
        public static dynamic CreateFillStyle(
            dynamic app,
            string name,
            int r, int g, int b,
            bool replaceExisting = true)
        {
            try
            {
                dynamic sheet = app.ActiveDocument.StyleSheet;
                dynamic st = sheet.CreateStyle(CategoryFill, string.Empty, name, replaceExisting);
                // cdrUniformFillStyle = 1
                st.Fill.Type = 1;
                st.Fill.PrimaryColor.RGBAssign(r, g, b);
                return st;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 创建一个 RGB 纯色轮廓样式。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.CreateStyle("outline", "", name, True)</code>
        /// 后设置 <c>Style.Outline.Width</c> 和 <c>Style.Outline.Color</c>。
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">样式名称。</param>
        /// <param name="widthMm">轮廓宽度（毫米）。</param>
        /// <param name="r">轮廓颜色 R（0–255）。</param>
        /// <param name="g">轮廓颜色 G（0–255）。</param>
        /// <param name="b">轮廓颜色 B（0–255）。</param>
        /// <param name="replaceExisting">同名样式存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>创建的 <c>Style</c> COM 对象；失败时返回 <c>null</c>。</returns>
        public static dynamic CreateOutlineStyle(
            dynamic app,
            string name,
            double widthMm,
            int r, int g, int b,
            bool replaceExisting = true)
        {
            try
            {
                dynamic sheet = app.ActiveDocument.StyleSheet;
                dynamic st = sheet.CreateStyle(CategoryOutline, string.Empty, name, replaceExisting);
                st.Outline.Width = widthMm;
                st.Outline.Color.RGBAssign(r, g, b);
                return st;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 创建一个字符样式（字体 / 字号 / 粗斜体）。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">样式名称。</param>
        /// <param name="fontName">字体名称（如 <c>"Arial"</c>、<c>"微软雅黑"</c>）。</param>
        /// <param name="sizePoints">字号（磅）。</param>
        /// <param name="bold">是否粗体。</param>
        /// <param name="italic">是否斜体。</param>
        /// <param name="replaceExisting">同名样式存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>创建的 <c>Style</c> COM 对象；失败时返回 <c>null</c>。</returns>
        public static dynamic CreateCharacterStyle(
            dynamic app,
            string name,
            string fontName,
            double sizePoints,
            bool bold = false,
            bool italic = false,
            bool replaceExisting = true)
        {
            try
            {
                dynamic sheet = app.ActiveDocument.StyleSheet;
                dynamic st = sheet.CreateStyle(CategoryCharacter, string.Empty, name, replaceExisting);
                st.Character.Font   = fontName;
                st.Character.Size   = sizePoints;
                st.Character.Bold   = bold;
                st.Character.Italic = italic;
                return st;
            }
            catch
            {
                return null;
            }
        }

        // ─────────────────── 样式集（多类别）创建 ────────────────────────────

        /// <summary>
        /// 通过临时形状创建一个包含填充 + 轮廓的样式集（StyleSet）。
        /// <para>
        /// <b>原理</b>：CorelDRAW 的 <c>CreateStyleSet</c> 返回空容器，
        /// 其 <c>.Fill</c>、<c>.Outline</c> 属性均为 <c>null</c>，无法直接赋值。
        /// 本方法绕过此限制：先在画面外创建临时形状并配置外观，
        /// 再通过 <c>CreateStyleSetFromShape</c> 捕获为样式集，最后删除临时形状。
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">样式集名称。</param>
        /// <param name="fillR">填充颜色 R（0–255）。</param>
        /// <param name="fillG">填充颜色 G（0–255）。</param>
        /// <param name="fillB">填充颜色 B（0–255）。</param>
        /// <param name="outlineWidthMm">轮廓宽度（毫米）；传 0 表示无轮廓。</param>
        /// <param name="outlineR">轮廓颜色 R（0–255）。</param>
        /// <param name="outlineG">轮廓颜色 G（0–255）。</param>
        /// <param name="outlineB">轮廓颜色 B（0–255）。</param>
        /// <param name="replaceExisting">同名样式集存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>
        /// 创建成功时返回对应的 <c>Style</c> COM 对象（<c>IsStyleSet = true</c>）；
        /// 失败时返回 <c>null</c>。
        /// </returns>
        /// <example>
        /// <code>
        /// // 创建蓝色填充 + 0.5 mm 黑色轮廓的样式集
        /// CorelStyleSheetHelper.CreateStyleSetViaShape(
        ///     app, "品牌图形",
        ///     fillR: 31, fillG: 73, fillB: 125,
        ///     outlineWidthMm: 0.5,
        ///     outlineR: 0, outlineG: 0, outlineB: 0);
        ///
        /// // 应用到形状
        /// CorelStyleSheetHelper.ApplyStyle(app.ActiveShape, "品牌图形");
        /// </code>
        /// </example>
        public static dynamic CreateStyleSetViaShape(
            dynamic app,
            string name,
            int fillR, int fillG, int fillB,
            double outlineWidthMm = 0.5,
            int outlineR = 0, int outlineG = 0, int outlineB = 0,
            bool replaceExisting = true)
        {
            dynamic tmpShape = null;
            try
            {
                // 在画面外创建临时矩形（-9999 坐标处），不影响正常内容
                tmpShape = app.ActiveLayer.CreateRectangle(-9999, -9999, -9998, -9998);

                // 设置填充
                dynamic fillColor = app.CreateRGBColor(fillR, fillG, fillB);
                tmpShape.Fill.ApplyUniformFill(fillColor);

                // 设置轮廓
                if (outlineWidthMm > 0)
                {
                    tmpShape.Outline.Width = outlineWidthMm;
                    tmpShape.Outline.Color.RGBAssign(outlineR, outlineG, outlineB);
                }
                else
                {
                    tmpShape.Outline.Width = 0;
                }

                // 从临时形状捕获为样式集
                dynamic createdStyles = app.ActiveDocument.StyleSheet
                    .CreateStyleSetFromShape(tmpShape, name, replaceExisting);

                // 返回新建的样式集对象（找到同名样式）
                return app.ActiveDocument.StyleSheet.FindStyle(name);
            }
            catch
            {
                return null;
            }
            finally
            {
                // 无论是否成功，都删除临时形状
                try { tmpShape?.Delete(); }
                catch { }
            }
        }

        /// <summary>
        /// 通过子样式方法创建一个包含填充 + 轮廓的样式集（StyleSet）。
        /// <para>
        /// <b>原理</b>：先调用 <c>CreateStyleSet</c> 创建空容器，
        /// 再分别调用 <c>styleSet.CreateStyle("fill")</c> 和
        /// <c>styleSet.CreateStyle("outline")</c> 激活对应属性，
        /// 然后通过子样式对象设置具体属性值。
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">样式集名称。</param>
        /// <param name="fillR">填充颜色 R（0–255）。</param>
        /// <param name="fillG">填充颜色 G（0–255）。</param>
        /// <param name="fillB">填充颜色 B（0–255）。</param>
        /// <param name="outlineWidthMm">轮廓宽度（毫米）；传 0 表示不添加轮廓子样式。</param>
        /// <param name="outlineR">轮廓颜色 R（0–255）。</param>
        /// <param name="outlineG">轮廓颜色 G（0–255）。</param>
        /// <param name="outlineB">轮廓颜色 B（0–255）。</param>
        /// <param name="replaceExisting">同名样式集存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>
        /// 创建成功时返回 <c>Style</c> COM 对象（<c>IsStyleSet = true</c>）；
        /// 失败时返回 <c>null</c>。
        /// </returns>
        public static dynamic CreateStyleSetViaSubStyles(
            dynamic app,
            string name,
            int fillR, int fillG, int fillB,
            double outlineWidthMm = 0.5,
            int outlineR = 0, int outlineG = 0, int outlineB = 0,
            bool replaceExisting = true)
        {
            try
            {
                dynamic sheet = app.ActiveDocument.StyleSheet;
                // 创建空样式集容器（此时 ss.Fill / ss.Outline 均为 null）
                dynamic ss = sheet.CreateStyleSet(string.Empty, name, replaceExisting);

                // 在容器内创建 fill 子样式 → 激活 ss.Fill，使其不再为 null
                dynamic fillSub = ss.CreateStyle(CategoryFill);
                fillSub.Fill.Type = 1; // cdrUniformFillStyle = 1
                fillSub.Fill.PrimaryColor.RGBAssign(fillR, fillG, fillB);

                // 在容器内创建 outline 子样式（可选）
                if (outlineWidthMm > 0)
                {
                    dynamic outlineSub = ss.CreateStyle(CategoryOutline);
                    outlineSub.Outline.Width = outlineWidthMm;
                    outlineSub.Outline.Color.RGBAssign(outlineR, outlineG, outlineB);
                }

                return ss;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 从现有形状捕获为样式集（StyleSet），一次性包含填充 + 轮廓 + 字符等所有属性。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.CreateStyleSetFromShape(Shape, Name, ReplaceExisting)</code>
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="shape">已配置好外观的 <c>Shape</c> COM 对象。</param>
        /// <param name="name">样式集名称。</param>
        /// <param name="replaceExisting">同名样式集存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>
        /// 创建成功时返回 <c>Style</c> COM 对象（<c>IsStyleSet = true</c>）；
        /// 失败时返回 <c>null</c>。
        /// </returns>
        public static dynamic CreateStyleSetFromShape(
            dynamic app,
            dynamic shape,
            string name,
            bool replaceExisting = true)
        {
            try
            {
                app.ActiveDocument.StyleSheet
                    .CreateStyleSetFromShape(shape, name, replaceExisting);
                return app.ActiveDocument.StyleSheet.FindStyle(name);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 从现有形状捕获单类别样式（如只捕获填充或只捕获轮廓）。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.CreateStyleFromShape(Shape, Category, Name, ReplaceExisting)</code>
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="shape">已配置好外观的 <c>Shape</c> COM 对象。</param>
        /// <param name="category">
        /// 类别名称，使用本类中的常量：
        /// <see cref="CategoryFill"/>、<see cref="CategoryOutline"/>、
        /// <see cref="CategoryCharacter"/>、<see cref="CategoryParagraph"/>、
        /// <see cref="CategoryFrame"/>。
        /// </param>
        /// <param name="name">样式名称。</param>
        /// <param name="replaceExisting">同名样式存在时是否覆盖，默认 <c>true</c>。</param>
        /// <returns>
        /// 创建成功时返回 <c>Style</c> COM 对象；
        /// 失败时返回 <c>null</c>。
        /// </returns>
        public static dynamic CreateStyleFromShape(
            dynamic app,
            dynamic shape,
            string category,
            string name,
            bool replaceExisting = true)
        {
            try
            {
                app.ActiveDocument.StyleSheet
                    .CreateStyleFromShape(shape, category, name, replaceExisting);
                return app.ActiveDocument.StyleSheet.FindStyle(name);
            }
            catch
            {
                return null;
            }
        }

        // ─────────────────────────── 查找与应用 ──────────────────────────────

        /// <summary>
        /// 按名称查找样式或样式集。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.FindStyle(Name)</code>
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">样式或样式集名称。</param>
        /// <returns>找到时返回 <c>Style</c> COM 对象；未找到时返回 <c>null</c>。</returns>
        public static dynamic FindStyle(dynamic app, string name)
        {
            try
            {
                return app.ActiveDocument.StyleSheet.FindStyle(name);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 将命名样式或样式集应用到形状。
        /// <para>
        /// 对应 VBA：<code>Shape.ApplyStyle(StyleName)</code>
        /// </para>
        /// </summary>
        /// <param name="shape"><c>Shape</c> COM 对象。</param>
        /// <param name="styleName">样式或样式集名称。</param>
        /// <returns><c>true</c> 表示应用成功；<c>false</c> 表示失败。</returns>
        public static bool ApplyStyle(dynamic shape, string styleName)
        {
            try
            {
                shape.ApplyStyle(styleName);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 对当前文档中的多个形状批量应用同一命名样式。
        /// </summary>
        /// <param name="shapes">
        /// <c>Shape</c> COM 对象的集合（可以是 <c>IEnumerable&lt;dynamic&gt;</c>
        /// 或 CorelDRAW 的 <c>ShapeRange</c> COM 枚举对象）。
        /// </param>
        /// <param name="styleName">样式或样式集名称。</param>
        /// <returns>成功应用的形状数量。</returns>
        public static int ApplyStyleToShapes(dynamic shapes, string styleName)
        {
            int count = 0;
            try
            {
                foreach (dynamic shape in shapes)
                {
                    try
                    {
                        shape.ApplyStyle(styleName);
                        count++;
                    }
                    catch { }
                }
            }
            catch { }
            return count;
        }

        // ────────────────────────── 样式管理 ─────────────────────────────────

        /// <summary>
        /// 删除指定名称的样式或样式集。
        /// <para>
        /// 对应 VBA：<code>StyleSheet.FindStyle(Name).Delete()</code>
        /// </para>
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="name">要删除的样式名称。</param>
        /// <returns><c>true</c> 表示删除成功；<c>false</c> 表示未找到或删除失败。</returns>
        public static bool DeleteStyle(dynamic app, string name)
        {
            try
            {
                dynamic st = app.ActiveDocument.StyleSheet.FindStyle(name);
                if (st == null) return false;
                return (bool)st.Delete();
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 重命名指定样式或样式集。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="oldName">当前样式名称。</param>
        /// <param name="newName">新样式名称。</param>
        /// <returns><c>true</c> 表示重命名成功；<c>false</c> 表示失败。</returns>
        public static bool RenameStyle(dynamic app, string oldName, string newName)
        {
            try
            {
                dynamic st = app.ActiveDocument.StyleSheet.FindStyle(oldName);
                if (st == null) return false;
                return (bool)st.Rename(newName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 获取当前文档中所有顶层样式（不含样式集）的名称列表。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <returns>样式名称列表（不含样式集）；失败时返回空列表。</returns>
        public static IReadOnlyList<string> GetStyleNames(dynamic app)
        {
            var result = new List<string>();
            try
            {
                foreach (dynamic st in app.ActiveDocument.StyleSheet.Styles)
                {
                    try { result.Add((string)st.Name); }
                    catch { }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// 获取当前文档中所有顶层样式集的名称列表。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <returns>样式集名称列表；失败时返回空列表。</returns>
        public static IReadOnlyList<string> GetStyleSetNames(dynamic app)
        {
            var result = new List<string>();
            try
            {
                foreach (dynamic st in app.ActiveDocument.StyleSheet.StyleSets)
                {
                    try { result.Add((string)st.Name); }
                    catch { }
                }
            }
            catch { }
            return result;
        }

        // ──────────────────────── 样式表持久化 ───────────────────────────────

        /// <summary>
        /// 将当前文档的样式和样式集导出为 <c>.cdss</c> 文件。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="filePath">导出文件的完整路径（建议扩展名 <c>.cdss</c>）。</param>
        /// <param name="includeStyles">是否包含普通样式（<c>IsStyleSet = false</c>），默认 <c>true</c>。</param>
        /// <param name="includeStyleSets">是否包含样式集（<c>IsStyleSet = true</c>），默认 <c>true</c>。</param>
        /// <returns><c>true</c> 表示导出成功；<c>false</c> 表示失败。</returns>
        public static bool ExportStyles(
            dynamic app,
            string filePath,
            bool includeStyles = true,
            bool includeStyleSets = true)
        {
            try
            {
                app.ActiveDocument.StyleSheet.Export(
                    filePath,
                    includeStyles,
                    includeStyleSets,
                    false);  // ObjectDefaults = false（不导出内置默认样式）
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 从 <c>.cdss</c> 文件导入样式和样式集到当前文档。
        /// </summary>
        /// <param name="app">CorelDRAW Application COM 对象。</param>
        /// <param name="filePath">要导入的文件完整路径。</param>
        /// <param name="mergeStyles">
        /// <c>true</c> 表示合并（同名样式保留旧版，新样式追加）；
        /// <c>false</c> 表示覆盖同名样式。
        /// </param>
        /// <returns><c>true</c> 表示导入成功；<c>false</c> 表示失败。</returns>
        public static bool ImportStyles(
            dynamic app,
            string filePath,
            bool mergeStyles = true)
        {
            try
            {
                app.ActiveDocument.StyleSheet.Import(
                    filePath,
                    mergeStyles,
                    true,   // Styles
                    true);  // StyleSets
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
