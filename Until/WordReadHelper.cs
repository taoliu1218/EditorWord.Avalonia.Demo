using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;

namespace EditWord.Avalonia.Until;

public static class WordReadHelper
{
    /// <summary>
    /// docx中的xml数据节点
    /// </summary>
    private static AvaloniaDictionary<string, ZipArchiveEntry> _zipArchive = new();

    /// <summary>
    /// 主要命名空间xmlns:w
    /// </summary>
    private static XNamespace _w = XNamespace.Get("");

    /// <summary>
    /// 渲染word文档
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    /// <exception cref="IOException"></exception>
    /// <exception cref="InvalidOperationException"></exception>
    public static Control RenderDocument(string filePath)
    {
        // 先创建一个父容器，模拟word文档的灰色区域
        var scrollView = new ScrollViewer()
        {
            Background = Brushes.Gray,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            VerticalScrollBarVisibility = ScrollBarVisibility.Hidden,
        };
        // 创建文档容器，以Border为例，设置为白色
        var border = new Border()
        {
            Background = Brushes.White,
            Padding = new Thickness(3),
            Margin = new Thickness(10, 50)
        };

        scrollView.Content = border;


        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("文件路径不能为空", nameof(filePath));

        if (!File.Exists(filePath))
            throw new IOException("文件不存在");

        var file = new FileInfo(filePath);
        var extension = file.Extension.ToLowerInvariant();

        // 判断是否是 Word 文件
        if (extension != ".docx")
            throw new InvalidOperationException("文件类型不正确，必须是 Word 文档 (.docx)");


        using var zip = ZipFile.OpenRead(filePath);
        // 将xml节点保存到字典中，以便后续加载子节点查找链接数据
        _zipArchive = new AvaloniaDictionary<string, ZipArchiveEntry>();
        foreach (var entry in zip.Entries.Where(p => !string.IsNullOrEmpty(p.Name)))
        {
            _zipArchive.Add(entry.Name, entry);
        }

        // 取得document文档节点
        var documentEntry = _zipArchive["document.xml"];
        if (documentEntry == null)
            throw new InvalidOperationException("无法在文档中找到 word/document.xml");


        #region 渲染Body

        //打开文档节点，先获取文档宽高，来设置整个page的宽高
        using var stream = documentEntry.Open();
        var doc = XDocument.Load(stream);
        //取得document文档的w命名空间
        var wAttr = doc.Root!.Attributes()
            .First(a => a is { IsNamespaceDeclaration: true, Name.LocalName: "w" });
        _w = XNamespace.Get(wAttr.Value);

        //取得pageSize节点获取宽高，换算成DPI像素宽高
        var size = doc.Descendants(_w + "body").First().Element(_w + "sectPr")!.Element(_w + "pgSz")!;
        int.TryParse(size.Attribute(_w + "h")!.Value, out var tipHeight);
        int.TryParse(size.Attribute(_w + "w")!.Value, out var tipWidth);

        var width = NumericalConversion(tipWidth);
        var height = NumericalConversion(tipHeight);
        border.Height = height;
        border.Width = width;

        var pageChildren = new Grid()
        {
            RowDefinitions = new RowDefinitions("Auto,*,Auto")
        };

        var bodyElement = doc.Root.Element(_w + "body");
        if (bodyElement != null)
        {
            var bodyControl = RenderDocumentControls(bodyElement);
            Grid.SetRow(bodyControl, 1);
            pageChildren.Children.Add(bodyControl);
        }

        #endregion

        #region 渲染页眉

        var headerEntry = zip.GetEntry("word/header1.xml");
        if (headerEntry != null)
        {
            using var headerStream = headerEntry.Open();
            var headerDoc = XDocument.Load(headerStream);
            if (headerDoc.Root != null)
            {
                var headerControl = RenderDocumentControls(headerDoc.Root);
                Grid.SetRow(headerControl, 0);
                pageChildren.Children.Add(headerControl);
            }
        }

        #endregion

        #region 渲染页脚

        var footerEntry = zip.GetEntry("word/footer1.xml");
        if (footerEntry != null)
        {
            using var footerStream = footerEntry.Open();
            var footerDoc = XDocument.Load(footerStream);
            if (footerDoc.Root != null)
            {
                var footerControl = RenderDocumentControls(footerDoc.Root);
                Grid.SetRow(footerControl, 2);
                pageChildren.Children.Add(footerControl);
            }
        }

        #endregion


        border.Child = pageChildren;

        return scrollView;
    }

    /// <summary>
    /// 渲染文档部件
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static Grid RenderDocumentControls(XElement element)
    {
        var grid = new Grid();
        // 筛选掉没有内容的段落
        var elements = element.Elements().Where(p => p.Descendants(_w + "r").Any()).ToList();


        if (elements.Count == 0) return grid;

        // 设置Grid的row，和文档节点数对应
        var rows = elements.Select(p => "Auto").ToList();
        var rowDefinitions = string.Join(",", rows);
        grid.RowDefinitions = new RowDefinitions(rowDefinitions);

        // 遍历生成控件
        for (var i = 0; i < elements.Count; i++)
        {
            var el = elements[i];

            //根据节点类型生成对应的控件
            var gridControl = el.Name.LocalName switch
            {
                "p" => LoadParagraph(el),
                "tbl" => LoadTable(el),
                _ => null
            };

            gridControl ??= new Border();

            Grid.SetRow(gridControl, i);
            grid.Children.Add(gridControl);
        }

        return grid;
    }

    /// <summary>
    /// 生成body文档区域控件
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    private static Grid RenderBodyControls(XDocument document)
    {
        var grid = new Grid();

        var bodyXmlElement = document.Root!.Element(_w + "body");
        if (bodyXmlElement == null)
        {
            throw new Exception("未读取到body内容");
        }

        var els = bodyXmlElement.Elements().ToList();

        // 筛选掉没有内容的段落
        var elements = bodyXmlElement.Elements().Where(p => p.Descendants(_w + "r").Any()).ToList();


        if (elements.Count == 0) return grid;

        // 设置Grid的row，和文档节点数对应
        var rows = elements.Select(p => "Auto").ToList();
        var rowDefinitions = string.Join(",", rows);
        grid.RowDefinitions = new RowDefinitions(rowDefinitions);

        // 遍历生成控件
        for (var i = 0; i < elements.Count; i++)
        {
            var el = elements[i];

            //根据节点类型生成对应的控件
            var gridControl = el.Name.LocalName switch
            {
                "p" => LoadParagraph(el),
                "tbl" => LoadTable(el),
                _ => null
            };

            gridControl ??= new Border();

            Grid.SetRow(gridControl, i);
            grid.Children.Add(gridControl);
        }

        return grid;
    }

    /// <summary>
    /// 加载表格
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static Control LoadTable(XElement element)
    {
        // 设置表格Grid布局 有多少个tableRow就有多少行
        var parentGrid = new Grid();
        var rows = element.Descendants(_w + "tr").ToList();
        var rowWidths = rows.Select(p => "Auto").ToList();
        parentGrid.RowDefinitions = new RowDefinitions(string.Join(",", rowWidths));
        Console.WriteLine($"表格总共：{rows.Count}行");

        var colStyle = element.Element(_w + "tblGrid")?.Elements().ToList();


        // 遍历行节点生成对应控件
        for (var i = 0; i < rows.Count; i++)
        {
            var row = rows[i];

            // 取得单元格设置
            ColumnDefinitions? columnDefinitions = null;
            if (colStyle != null)
            {
                var colWidthValue = colStyle.Select(p => Convert.ToInt32(p.Attribute(_w + "w").Value)).ToList();

                var widths = colWidthValue.Select(NumericalConversion).Select(p => $"{p}").ToList();

                Console.WriteLine($"第{i + 1}行总共有{widths.Count}列");

                columnDefinitions = new ColumnDefinitions(string.Join(",", widths));
            }

            //创建行数据Grid，并设置列
            var rowGrid = new Grid();
            if (columnDefinitions != null)
            {
                rowGrid.ColumnDefinitions = columnDefinitions;
            }

            // 设置行最低高度
            var trHeightEl = row.Element(_w + "trPr")?.Element(_w + "trHeight");
            if (trHeightEl != null)
            {
                if (int.TryParse(trHeightEl.Attribute(_w + "val")?.Value, out var hVal))
                {
                    rowGrid.MinHeight = NumericalConversion(hVal);
                }
            }

            //取得行所有单元格节点
            var tableCellElements = row.Elements(_w + "tc").ToList();

            //遍历单元格节点集合，生成控件
            for (int j = 0; j < tableCellElements.Count; j++)
            {
                var tableCellElement = tableCellElements[j];
                //创建单元格边框
                var tableCellBorder = new Border()
                {
                    BorderThickness = new Thickness(1),
                    BorderBrush = Brushes.Gray
                };

                // 设置单元格边框
                var tcBorderElement = tableCellElement.Descendants(_w + "tcBorders").FirstOrDefault();
                if (tcBorderElement != null)
                {
                    var isNoBorder = tcBorderElement.Elements().Any(p => p.Attribute(_w + "val").Value == "nil");
                    if (isNoBorder)
                    {
                        tableCellBorder.BorderThickness = new Thickness(0);
                        tableCellBorder.BorderBrush = Brushes.Transparent;
                    }
                }
                else
                {
                    tableCellBorder.BorderThickness = new Thickness(1);
                }

                var pElement = tableCellElement.Element(_w + "p");

                if (pElement != null)
                {
                    var control = LoadParagraph(pElement);
                    tableCellBorder.Child = control;
                }

                Grid.SetColumn(tableCellBorder, j);

                var colSpan = tableCellElement.Descendants(_w + "gridSpan").FirstOrDefault()?.Attribute(_w + "val")
                    ?.Value;
                if (!string.IsNullOrEmpty(colSpan))
                {
                    int val = Convert.ToInt32(colSpan);
                    Grid.SetColumnSpan(tableCellBorder, val);
                    j = val - 1;
                }


                rowGrid.Children.Add(tableCellBorder);
                Console.WriteLine($"第{i + 1}行第{j + 1}列设置成功");
            }

            //设置布局
            Grid.SetRow(rowGrid, i);
            parentGrid.Children.Add(rowGrid);
            Console.WriteLine($"第{i + 1}行设置成功");
        }


        return parentGrid;
    }


    /// <summary>
    /// 加载段落
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static Control LoadParagraph(XElement element)
    {
        var panel = new Grid();

        // 筛选不需要的节点，pPr为段落属性节点，它定义了该段落（<w:p>）的所有样式、对齐方式、缩进、行距、边距、编号、样式引用等。这里暂时不做处理
        // <w:r> 是 文字运行节点（Run），是 Word 文档中最核心的文字单位。代表 一段连续、样式一致的文字或内容片段。
        // <w:rPr> Run Properties（文字属性），控制字体、字号、加粗、斜体、颜色、下划线等
        // <w:t> Text（文本内容），存放文字内容
        // <w:drawing> Drawing 对象（图片、图形、图表）
        // <w:br> 换行符（Break）
         var items = element.Elements().Where(p => p.Name.LocalName != "pPr").ToList();

        //var items = element.Descendants(_w + "r").ToList();

        // 这个段落是否是图片
        var isImage = element.Descendants(_w + "drawing").FirstOrDefault() != null;

        //是图片的话进行图片的渲染
        if (isImage)
        {
            var drawingList = element.Descendants(_w + "drawing").ToList();
            foreach (var img in drawingList.Select(LoadImage))
            {
                panel.Children.Add(img);
            }
        }
        //不是则进行文本的渲染
        else
        {
            var style = element.Descendants(_w + "jc").FirstOrDefault()?.Attribute(_w + "val")?.Value;

            #region 如果整个段落只有一个文本节点，则直接渲染一个textbox

            if (items.Count == 1)
            {
                var txt = InitNewTextBox();

                if (style == "center")
                {
                    txt.HorizontalContentAlignment = HorizontalAlignment.Center;
                }

                txt.Text = items.First().Value;
                return txt;
            }

            #endregion


            #region 如果整个段落只有纯文本且无BookMark则渲染单文本框

            var isSingleText = true;
            foreach (var item in items)
            {
                if (!item.Elements(_w + "t").Any() || item.Descendants(_w + "bookmarkStart").Any())
                {
                    isSingleText = false;
                    break;
                }
            }

            if (isSingleText)
            {
                var txt = InitNewTextBox();
                txt.Text = string.Join("", items.Elements(_w + "t").Select(p => p.Value));
                return txt;
            }

            #endregion


            #region 段落有多个节点的时候渲染不同的控件

            while (items.Count != 0)
            {
                var firstElement = items.First();
                var elContent = string.Empty;

                var txt = InitNewTextBox();
                if (style == "center")
                {
                    txt.HorizontalContentAlignment = HorizontalAlignment.Center;
                }

                if (firstElement.Name == _w + "r")
                {
                    elContent += firstElement.Value;
                    var nextElement = firstElement.NextNode as XElement;
                    while (nextElement != null && nextElement.Name == _w + "r")
                    {
                        elContent += nextElement.Value;
                        nextElement = nextElement.NextNode as XElement;
                    }


                    txt.Text = elContent;

                    elContent = string.Empty;
                    panel.Children.Add(txt);
                }

                else if (firstElement.Name == _w + "bookmarkStart" &&
                         firstElement.Attribute(_w + "name")!.Value != "_GoBack")
                {
                    var bmId = firstElement.Attribute(_w + "id")!.Value;
                    var bName = firstElement.Attribute(_w + "name")!.Value;
                    var collecting = false;
                    foreach (var node in element.Elements())
                    {
                        if (node.Name == _w + "bookmarkStart" && node.Attribute(_w + "id")?.Value == bmId)
                            collecting = true;

                        if (!collecting) continue;

                        items.Remove(node);
                        elContent += node.Value;
                        if (node.Name == _w + "bookmarkEnd" && node.Attribute(_w + "id")?.Value == bmId)
                        {
                            break;
                        }
                    }

                    txt[!TextBox.TextProperty] = new Binding(bName);
                    txt.Text = elContent;
                    txt.HorizontalAlignment = HorizontalAlignment.Stretch;
                    txt.Name = $"WordEditor_{bName}";

                    panel.Children.Add(txt);
                }

                items.Remove(firstElement);
            }

            #endregion
        }

        return panel;
    }

    /// <summary>
    /// 加载图片
    /// </summary>
    /// <param name="element"></param>
    /// <returns></returns>
    private static Image LoadImage(XElement element)
    {
        var img = new Image();

        // 取得wp命名空间，用来获取图片宽高和 a命名空间
        var wpAttribute = element.Document!.Root!.Attributes()
            .First(a => a is { IsNamespaceDeclaration: true, Name.LocalName: "wp" });
        var wpNameSpace = XNamespace.Get(wpAttribute.Value);

        var graphEl = element.Descendants().Where(p => p.Name.LocalName == "graphicFrameLocks");

        // 取得a命名空间，用来取得图片链接信息
        var aNameAttribute =
            graphEl.Attributes().First(a => a is { IsNamespaceDeclaration: true, Name.LocalName: "a" });
        var aNameSpace = XNamespace.Get(aNameAttribute.Value);


        // 取得图片宽高，换算成dpi像素宽高
        var size = element.Descendants(wpNameSpace + "extent").First();
        var emuWidth = size.Attribute("cx")!.Value;
        var emuHeight = size.Attribute("cy")!.Value;

        var imgWidth = Convert.ToDouble(emuWidth) / 914400 * 96;
        var imgHeight = Convert.ToDouble(emuHeight) / 914400 * 96;

        img.Width = imgWidth;
        img.Height = imgHeight;

        // 取得<a:blip>节点，这里保存了图片链接节点的Id，指向document.xml.rels下的Relationship对应Id节点
        var imageLinkEl = element.Descendants(aNameSpace + "blip").First();

        // 取得r命名空间，用来获取<a:blip>上的Attribute为r:embed的值，这个值为链接节点Id
        var rNameAttribute = element.Document.Root.Attributes()
            .First(a => a is { IsNamespaceDeclaration: true, Name.LocalName: "r" });
        var rNameSpace = XNamespace.Get(rNameAttribute.Value);

        // 节点链接Rid
        var linkId = imageLinkEl.Attribute(rNameSpace + "embed")!.Value;

        //文档rule节点数据
        var rules = _zipArchive["document.xml.rels"];

        using var stream = rules.Open();
        var doc = XDocument.Load(stream);
        var ruleElement = doc.Root!.Descendants().ToList();

        // 取得链接节点数据
        var linkRuleEl = ruleElement.First(p => p.Attribute("Id")!.Value == linkId);

        // 图片链接xml节点名称
        var imageLinkXmlName = linkRuleEl.Attribute("Target")!.Value;
        // 图片压缩数据节点key
        var imageXmlKey = imageLinkXmlName.Split("/").Last();

        // 将图片二进制数据转换为BitMap
        var imageZip = _zipArchive[imageXmlKey];
        using var imgStream = imageZip.Open();
        using var ms = new MemoryStream();
        imgStream.CopyTo(ms);

        var base64String = Convert.ToBase64String(ms.ToArray());

        byte[] imageBytes = Convert.FromBase64String(base64String);

        // 转换为 Avalonia Bitmap
        using var imgMs = new MemoryStream(imageBytes);
        var bitMap = new Bitmap(imgMs);

        img.Source = bitMap;

        return img;
    }

    // Twip转换Dpi像素换算值
    private const double TwipToDip = 96.0 / 1440.0;

    /// <summary>
    /// 初始化输入框控件
    /// </summary>
    /// <returns></returns>
    private static TextBox InitNewTextBox()
    {
        var txt = new TextBox()
        {
            BorderThickness = new Thickness(0),
            TextWrapping = TextWrapping.Wrap,
            AcceptsReturn = true,
            FontSize = 18,
            Classes = { "NoShadow" },
            [!TextBox.IsReadOnlyProperty] = new Binding("IsReadOnly")
        };
        txt.SetValue(ThemeVariantScope.ActualThemeVariantProperty, ThemeVariant.Light);

        return txt;
    }

    /// <summary>
    /// 宽高数值转换
    /// </summary>
    /// <param name="number"></param>
    /// <returns></returns>
    private static double NumericalConversion(int number)
    {
        return number * TwipToDip;
    }
}