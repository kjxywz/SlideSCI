using Markdig;  // 修改为使用签名版本的命名空间
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Text;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Collections.Generic; // Add this line

namespace Achuan的PPT插件
{
    public partial class Ribbon1
    {
        PowerPoint.Application app;
        private float copiedWidth;
        private float copiedHeight;
        private float copiedLeft;
        private float copiedTop;
        private bool isDarkBackground = true;  // Changed from false to true
        private float cropLeft;
        private float cropRight;
        private float cropTop;
        private float cropBottom;
        private bool hasCopiedCrop = false;
        private float originalHeight; // 添加变量存储原始图片高度
        private float currentCropedHeight;
        private const string CURRENT_VERSION = "1.0.0"; // Change this to match your current version
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void AddTitleToImage(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                float fontSize = float.Parse(fontSizeEditBox.Text);
                float distanceFromBottom = float.Parse(distanceFromBottomEditBox.Text);
                bool autoGroup = autoGroupCheckBox.Checked;
                string fontName = fontNameEditBox.Text;
                string titleText = titleTextEditBox.Text;

                foreach (PowerPoint.Shape selectedShape in sel.ShapeRange)
                {
                    if (selectedShape.Type == Office.MsoShapeType.msoPicture)
                    {
                        PowerPoint.Shape titleShape = slide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            selectedShape.Left,
                            selectedShape.Top + selectedShape.Height + distanceFromBottom,
                            selectedShape.Width,
                            fontSize * 2);

                        titleShape.TextFrame.TextRange.Text = titleText;
                        titleShape.TextFrame.TextRange.Font.Size = fontSize;
                        titleShape.TextFrame.TextRange.Font.NameFarEast = fontName; // Ensure FarEast font is set
                        titleShape.TextFrame.TextRange.Font.Name = fontName; // Ensure font is set
                        titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                        if (autoGroup)
                        {
                            PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new string[] { selectedShape.Name, titleShape.Name });
                            shapeRange.Group();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select an image to add a title.");
            }
        }

        private void pasteImgWidthHeight_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedWidth <= 0 || copiedHeight <= 0)
            {
                MessageBox.Show("Invalid copied dimensions. Please copy the dimensions again.");
                return;
            }

            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    shape.Width = copiedWidth;
                    shape.Height = copiedHeight;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste dimensions.");
            }
        }

        private void copyImgWidthHeight_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                copiedWidth = shape.Width;
                copiedHeight = shape.Height;
                // MessageBox.Show("Image dimensions copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy dimensions.");
            }
        }

        private void copyImgWidth_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                copiedWidth = shape.Width;
                // MessageBox.Show("Image width copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy width.");
            }
        }

        private void pasteImgWidth_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedWidth <= 0)
            {
                MessageBox.Show("Invalid copied width. Please copy the width again.");
                return;
            }

            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    shape.Width = copiedWidth;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste width.");
            }
        }

        private void copyImgHeight_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                copiedHeight = shape.Height;
                // MessageBox.Show("Image height copied!");
            }
            else
            {
                MessageBox.Show("Please select an image to copy height.");
            }
        }

        private void pasteImgHeight_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedHeight <= 0)
            {
                MessageBox.Show("Invalid copied height. Please copy the height again.");
                return;
            }

            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    shape.Height = copiedHeight;
                }
            }
            else
            {
                MessageBox.Show("Please select an image to paste height.");
            }
        }

        private void copyPosition_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                copiedLeft = shape.Left;
                copiedTop = shape.Top;
                // MessageBox.Show("Position copied!");
            }
            else
            {
                MessageBox.Show("Please select a shape to copy position.");
            }
        }

        private void pastePosition_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    shape.Left = copiedLeft;
                    shape.Top = copiedTop;
                }
            }
            else
            {
                MessageBox.Show("Please select a shape to paste position.");
            }
        }

        private void imgAutoAlign_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                int colNum;
                float colSpace;
                float rowSpace;
                float imgWidth = 0;
                float imgHeight = 0;

                if (!int.TryParse(imgAutoAlign_colNum.Text, out colNum) || colNum <= 0)
                {
                    MessageBox.Show("请输入有效的列数量。");
                    return;
                }

                if (!float.TryParse(imgAutoAlign_colSpace.Text, out colSpace) || colSpace < 0)
                {
                    MessageBox.Show("请输入有效的列间距。");
                    return;
                }

                if (!float.TryParse(imgAutoAlign_rowSpace.Text, out rowSpace) || rowSpace < 0)
                {
                    rowSpace = colSpace; // Use column spacing if row spacing is not provided
                }

                bool useCustomWidth = float.TryParse(imgWidthEditBpx.Text, out imgWidth) && imgWidth > 0;
                bool useCustomHeight = float.TryParse(imgHeightEditBox.Text, out imgHeight) && imgHeight > 0;

                PowerPoint.Shape firstShape = sel.ShapeRange[1];

                float startX = firstShape.Left;
                float startY = firstShape.Top;
                float currentX = startX;
                float currentY = startY;
                int currentCol = 0;

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (!useCustomHeight && !useCustomWidth)
                    {
                        shape.Height = firstShape.Height;
                    }
                    else
                    {
                        if (useCustomWidth)
                        {
                            shape.Width = imgWidth;
                        }
                        if (useCustomHeight)
                        {
                            shape.Height = imgHeight;
                        }
                    }


                    shape.Left = currentX;
                    shape.Top = currentY;

                    currentCol++;
                    if (currentCol >= colNum)
                    {
                        currentCol = 0;
                        currentX = startX; // Reset X position to startX
                        currentY += shape.Height + rowSpace;
                    }
                    else
                    {
                        currentX += shape.Width + colSpace;
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择要对齐的图片。");
            }
        }

        private void gallery1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("开发者: Achuan-2\n邮箱: achuan-2@outlook.com\nGithub地址：https://github.com/Achuan-2", "关于开发者");
        }

        private void insertCodeBlockButton_Click(object sender, RibbonControlEventArgs e)
        {

            // Create and configure input dialog
            Form inputDialog = new Form()
            {
                Width = 600,
                Height = 400,
                Text = "插入代码块",
                StartPosition = FormStartPosition.CenterScreen // Center the dialog on the screen
            };

            TextBox codeInput = new TextBox()
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 12)
            };

            ComboBox languageSelect = new ComboBox()
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // Add common programming languages
            languageSelect.Items.AddRange(new string[] {
                 "python", "matlab", "javascript",  "html", "css"
            });
            languageSelect.SelectedIndex = 0;

            Button okButton = new Button()
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom
            };

            // Add controls to form
            inputDialog.Controls.AddRange(new Control[] { codeInput, languageSelect, okButton });

            // Show dialog and process result
            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                string code = codeInput.Text.Trim();
                string language = languageSelect.SelectedItem.ToString();

                if (!string.IsNullOrEmpty(code))
                {
                    PowerPoint.Application app = Globals.ThisAddIn.Application;
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                    PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        100, 100, 500, 300);

                    // Set code block style
                    textBox.Fill.Solid();
                    textBox.Fill.ForeColor.RGB = isDarkBackground ?
                        ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                        ColorTranslator.ToOle(Color.White);
                    textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
                    textBox.Line.Weight = 1;

                    // Set the code without language markers
                    textBox.TextFrame.TextRange.Text = code;

                    // Apply base formatting
                    textBox.TextFrame.TextRange.Font.Name = "Consolas";
                    textBox.TextFrame.TextRange.Font.Size = 12;
                    textBox.TextFrame.TextRange.Font.Color.RGB = isDarkBackground ?
                        ColorTranslator.ToOle(Color.White) :
                        ColorTranslator.ToOle(Color.Black);
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                        PowerPoint.PpParagraphAlignment.ppAlignLeft;

                    // Set margins
                    textBox.TextFrame.MarginLeft = 10;
                    textBox.TextFrame.MarginRight = 10;
                    textBox.TextFrame.MarginTop = 5;
                    textBox.TextFrame.MarginBottom = 5;

                    // Apply syntax highlighting
                    var highlighter = new CodeHighlighter(isDarkBackground);
                    highlighter.ApplyHighlighting(textBox, code, language);

                    // Auto-size the textbox to fit content
                    textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
        }

        private void toggleBackgroundButton_Click(object sender, RibbonControlEventArgs e)
        {
            isDarkBackground = toggleBackgroundButton.Checked;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        // Update background color
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = isDarkBackground ?
                            ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                            ColorTranslator.ToOle(Color.White);

                        // Update text color
                        shape.TextFrame.TextRange.Font.Color.RGB = isDarkBackground ?
                            ColorTranslator.ToOle(Color.White) :
                            ColorTranslator.ToOle(Color.Black);
                    }
                }
            }
        }

        private void insertEquationButton_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;


            // Prompt user for LaTeX input
            Form inputDialog = new Form()
            {
                Width = 500,
                Height = 500,
                Text = "输入LaTeX公式",
                StartPosition = FormStartPosition.CenterScreen // Center the dialog on the screen
            };

            TextBox latexInputBox = new TextBox()
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 12)
            };

            Button okButton = new Button()
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Dock = DockStyle.Bottom
            };

            inputDialog.Controls.Add(latexInputBox);
            inputDialog.Controls.Add(okButton);

            if (inputDialog.ShowDialog() == DialogResult.OK)
            {
                string latexInput = latexInputBox.Text.Trim();

                // Remove surrounding $...$, $$...$$, \(...\), \[...\]
                if (latexInput.StartsWith("$") && latexInput.EndsWith("$"))
                {
                    latexInput = latexInput.Trim('$');
                }
                else if (latexInput.StartsWith("$$") && latexInput.EndsWith("$$"))
                {
                    latexInput = latexInput.Trim('$');
                }
                else if (latexInput.StartsWith(@"\(") && latexInput.EndsWith(@"\)"))
                {
                    latexInput = latexInput.Substring(2, latexInput.Length - 4);
                }
                else if (latexInput.StartsWith(@"\[") && latexInput.EndsWith(@"\]"))
                {
                    latexInput = latexInput.Substring(2, latexInput.Length - 4);
                }

                latexInput = latexInput.Replace("\r", "").Replace("\n", ""); // Remove line breaks

                if (!string.IsNullOrEmpty(latexInput))
                {
                    try
                    {
                        // Insert a new textbox in the center of the slide
                        PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        slide.Master.Width / 2 - 100, slide.Master.Height / 2 - 50, 500, 500);

                        // Select the newly inserted textbox
                        textBox.Select();
                        app.ActiveWindow.Selection.TextRange.Select();

                        // Run SwitchLatex
                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        PowerPoint.Shape equationShape = app.ActiveWindow.Selection.ShapeRange[1];
                        equationShape.TextFrame.TextRange.Characters(1, equationShape.TextFrame.TextRange.Text.Length - 1).Text = "\u24C9";

                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        app.ActiveWindow.Selection.TextRange.Select();
                        PowerPoint.Shape equationShape2 = app.ActiveWindow.Selection.ShapeRange[1];
                        // Set the LaTeX input to the equation shape
                        equationShape2.TextFrame.TextRange.Characters(1, equationShape2.TextFrame.TextRange.Text.Length - 1).Text = latexInput;

                        // Convert to professional format
                        app.CommandBars.ExecuteMso("EquationProfessional");

                        textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("An error occurred: " + ex.Message);
                    }
                }
            }
        }

        private int GetActualPosition(string text, int position)
        {
            return position - text.Substring(0, position).Count(c => c == '\r');
        }

        private string ConvertMarkdownToHtml(string markdown)
        {
            try
            {
                var html = Markdown.ToHtml(markdown);
                //MessageBox.Show($"Markdown转换: {html}");
                return html;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Markdown转换错误: {ex.Message}");
                return markdown; // 转换失败时返回原文本
            }
        }

        private void insertMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Form inputDialog = new Form
                {
                    Width = 600,
                    Height = 400,
                    Text = "插入Markdown",
                    StartPosition = FormStartPosition.CenterScreen
                };

                TextBox markdownInput = new TextBox
                {
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Fill,
                    Font = new Font("Consolas", 12)
                };

                Button okButton = new Button
                {
                    Text = "确定",
                    DialogResult = DialogResult.OK,
                    Dock = DockStyle.Bottom
                };

                inputDialog.Controls.Add(markdownInput);
                inputDialog.Controls.Add(okButton);

                DialogResult result = inputDialog.ShowDialog();

                if (result == DialogResult.OK)
                {
                    string markdown = markdownInput.Text?.Trim() ?? "";
                    if (!string.IsNullOrEmpty(markdown))
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                        // Split markdown into segments
                        var segments = SplitMarkdownIntoSegments(markdown);

                        float currentTop = slide.Master.Height / 3;  // Starting position
                        float left = (slide.Master.Width - 500) / 2; // Center horizontally

                        foreach (var segment in segments)
                        {
                            //MessageBox.Show($"Markdown分段: {segment.Content}");
                            if (segment.IsCodeBlock)
                            {
                                // Insert code block
                                PowerPoint.Shape codeShape = InsertCodeBlock(segment.Content, segment.Language, left, currentTop);
                                currentTop += codeShape.Height + 10;
                            }
                            else
                            {
                                // Process regular markdown text
                                string html = ProcessMarkdown(segment.Content);
                                if (!string.IsNullOrEmpty(html))
                                {
                                    CopyHtmlToClipBoard(segment.Content, html);
                                    PowerPoint.ShapeRange textContent = slide.Shapes.Paste();
                                    if (textContent != null && textContent.Count > 0)
                                    {
                                        PowerPoint.Shape textShape = textContent[1];
                                        textShape.Width = 500;
                                        textShape.Left = left;
                                        textShape.Top = currentTop;
                                        currentTop += textShape.Height + 10;
                                    }
                                }
                            }
                        }
                    }
                }

                inputDialog.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作过程中出错: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private class MarkdownSegment
        {
            public string Content { get; set; }
            public bool IsCodeBlock { get; set; }
            public string Language { get; set; }
        }

        private List<MarkdownSegment> SplitMarkdownIntoSegments(string markdown)
        {
            var segments = new List<MarkdownSegment>();
            var codeBlockRegex = new System.Text.RegularExpressions.Regex(
                @"```(\w*)\r?\n(.*?)\r?\n```",
                System.Text.RegularExpressions.RegexOptions.Singleline
            );

            int currentPosition = 0;
            var matches = codeBlockRegex.Matches(markdown);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                // Add text before code block if exists
                if (match.Index > currentPosition)
                {
                    string textBefore = markdown.Substring(currentPosition, match.Index - currentPosition);
                    if (!string.IsNullOrWhiteSpace(textBefore))
                    {
                        segments.Add(new MarkdownSegment
                        {
                            Content = textBefore.Trim(),
                            IsCodeBlock = false
                        });
                    }
                }

                // Add code block
                segments.Add(new MarkdownSegment
                {
                    Content = match.Groups[2].Value,
                    Language = string.IsNullOrEmpty(match.Groups[1].Value) ? "text" : match.Groups[1].Value,
                    IsCodeBlock = true
                });

                currentPosition = match.Index + match.Length;
            }

            // Add remaining text if exists
            if (currentPosition < markdown.Length)
            {
                string remainingText = markdown.Substring(currentPosition);
                if (!string.IsNullOrWhiteSpace(remainingText))
                {
                    segments.Add(new MarkdownSegment
                    {
                        Content = remainingText.Trim(),
                        IsCodeBlock = false
                    });
                }
            }

            return segments;
        }

        // Add this helper method for inserting code blocks
        private PowerPoint.Shape InsertCodeBlock(string code, string language, float left, float top)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 300);

            // Set code block style
            textBox.Fill.Solid();
            textBox.Fill.ForeColor.RGB = isDarkBackground ?
                ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                ColorTranslator.ToOle(Color.White);
            textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            textBox.Line.Weight = 1;

            textBox.TextFrame.TextRange.Text = code;

            // Apply base formatting
            textBox.TextFrame.TextRange.Font.Name = "Consolas";
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.TextFrame.TextRange.Font.Color.RGB = isDarkBackground ?
                ColorTranslator.ToOle(Color.White) :
                ColorTranslator.ToOle(Color.Black);
            textBox.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignLeft;

            // Set margins
            textBox.TextFrame.MarginLeft = 10;
            textBox.TextFrame.MarginRight = 10;
            textBox.TextFrame.MarginTop = 5;
            textBox.TextFrame.MarginBottom = 5;

            // Apply syntax highlighting
            var highlighter = new CodeHighlighter(isDarkBackground);
            highlighter.ApplyHighlighting(textBox, code, language);

            // Auto-size the textbox to fit content
            textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;

            return textBox;
        }

        private string ProcessMarkdown(string markdown)
        {
            // Process tables - add newline before tables and set styling
            markdown = System.Text.RegularExpressions.Regex.Replace(
                markdown,
                @"(^\||[^\\n]\|)",
                m => m.Value.StartsWith("\n") ? m.Value : "\n" + m.Value
            );

            // Remove code blocks
            var codeBlockRegex = new System.Text.RegularExpressions.Regex(
                @"```.*?\r?\n(.*?)\r?\n```",
                System.Text.RegularExpressions.RegexOptions.Singleline
            );

            markdown = codeBlockRegex.Replace(markdown, string.Empty);

            // Convert remaining markdown to HTML
            string html = Markdown.ToHtml(markdown);

            // Add table styling
            html = html.Replace("<table>", "<table style='width:500px; border-collapse:collapse; border:1px solid black;'>");
            html = html.Replace("<td>", "<td style='border:1px solid black; padding:5px;'>");
            html = html.Replace("<th>", "<th style='border:1px solid black; padding:5px;'>");

            html = html.Replace("<li>", "<li style='margin-left: 10px;'>");
            // 优化行内代码粘贴
            // <code>...</code> -> <span style='color: #C00000; font-family: Consolas;'>...</span>
            html = html.Replace("<code>", "<span style='color: #C00000; font-family: Consolas;'>");
            html = html.Replace("</code>", "</span>");

            // 设置
            html = $"<div style='font-family: 微软雅黑;'>{html}</div>";
            return html;
        }

        public void CopyHtmlToClipBoard(string markdown, string html)
        {
            var utf = Encoding.UTF8;
            var format = "Version:0.9\r\nStartHTML:{0:000000}\r\nEndHTML:{1:000000}\r\nStartFragment:{2:000000}\r\nEndFragment:{3:000000}\r\n";
            var text = "<html>\r\n<head>\r\n<meta http-equiv=\"Content-Type\" content=\"text/html; charset=" + utf.WebName + "\">\r\n<title>HTML clipboard</title>\r\n</head>\r\n<body>\r\n<!--StartFragment-->";
            var text2 = "<!--EndFragment-->\r\n</body>\r\n</html>\r\n";
            var s = string.Format(format, 0, 0, 0, 0);
            var byteCount = utf.GetByteCount(s);
            var byteCount2 = utf.GetByteCount(text);
            var byteCount3 = utf.GetByteCount(html);
            var byteCount4 = utf.GetByteCount(text2);
            var s2 = string.Format(format, byteCount, byteCount + byteCount2 + byteCount3 + byteCount4, byteCount + byteCount2, byteCount + byteCount2 + byteCount3) + text + html + text2;
            var dataObject = new DataObject();
            dataObject.SetData(DataFormats.Html, new MemoryStream(utf.GetBytes(s2)));
            dataObject.SetData(DataFormats.UnicodeText, markdown);
            Clipboard.SetDataObject(dataObject);
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://markdown.com.cn/editor");
        }

        private void copyCrop_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];

                // 保存裁剪设置
                cropLeft = shape.PictureFormat.CropLeft;
                cropRight = shape.PictureFormat.CropRight;
                cropTop = shape.PictureFormat.CropTop;
                cropBottom = shape.PictureFormat.CropBottom;

                // 保存原始高度
                currentCropedHeight = shape.Height;
                float croppedPixels = cropTop + cropBottom;
                originalHeight = currentCropedHeight + croppedPixels;

                hasCopiedCrop = true;
                //MessageBox.Show("已复制图片裁剪设置");
            }
            else
            {
                MessageBox.Show("请选择一个图片对象");
            }

        }

        private void pasteCrop_Click(object sender, RibbonControlEventArgs e)
        {
            if (!hasCopiedCrop)
            {
                MessageBox.Show("请先复制图片裁剪设置");
                return;
            }
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    try
                    {
                        // Store original position
                        float originalLeft = shape.Left;
                        float originalTop = shape.Top;

                        // Clear existing crop settings
                        shape.PictureFormat.CropLeft = 0;
                        shape.PictureFormat.CropRight = 0;
                        shape.PictureFormat.CropTop = 0;
                        shape.PictureFormat.CropBottom = 0;

                        // Restore to original height
                        shape.Height = originalHeight;

                        // Apply crop settings
                        shape.PictureFormat.CropLeft = cropLeft;
                        shape.PictureFormat.CropRight = cropRight;
                        shape.PictureFormat.CropTop = cropTop;
                        shape.PictureFormat.CropBottom = cropBottom;

                        shape.Height = currentCropedHeight;

                        // Restore original position
                        shape.Left = originalLeft;
                        shape.Top = originalTop;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"应用裁剪设置时出错: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("请选择要应用裁剪设置的图片");
            }
        }

        private void openGithub_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/Achuan-2/my_ppt_plugin/");
        }

    }

}

