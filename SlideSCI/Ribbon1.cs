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



namespace SlideSCI
{

    public partial class Ribbon1
    {
        PowerPoint.Application app;
        private float copiedWidth;
        private float copiedHeight;

        private List<float> copiedLeft = new List<float>();
        private List<float> copiedTop = new List<float>();

        private float cropLeft;
        private float cropRight;
        private float cropTop;
        private float cropBottom;
        private bool hasCopiedCrop = false;
        private float originalHeight; // 添加变量存储原始图片高度
        private float currentCropedHeight;



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;

            // Load Image Title Settings
            fontNameEditBox.Text = Properties.Settings.Default.TitleFontName;
            fontSizeEditBox.Text = Properties.Settings.Default.TitleFontSize;
            distanceFromBottomEditBox.Text = Properties.Settings.Default.TitleDistanceFromBottom;
            titleTextEditBox.Text = Properties.Settings.Default.TitleText;
            autoGroupCheckBox.Checked = Properties.Settings.Default.AutoGroup;

            // Load Image Label Settings
            labelOffsetXEditBox.Text = Properties.Settings.Default.LabelOffsetX;
            labelOffsetYEditBox.Text = Properties.Settings.Default.LabelOffsetY;
            labelTemplateComboBox.Text = Properties.Settings.Default.LabelTemplate;
            labelFontNameEditBox.Text = Properties.Settings.Default.LabelFontName;
            labelFontSizeEditBox.Text = Properties.Settings.Default.LabelFontSize;

            // Load Image Auto Align Settings
            imgAutoAlignSortTypeDropDown.SelectedItemIndex = Properties.Settings.Default.imgAutoAlignSortType;
            imgAutoAlign_colNum.Text = Properties.Settings.Default.ColNum;
            imgAutoAlign_colSpace.Text = Properties.Settings.Default.ColSpace;
            imgAutoAlign_rowSpace.Text = Properties.Settings.Default.RowSpace;
            imgWidthEditBpx.Text = Properties.Settings.Default.ImgWidth;
            imgHeightEditBox.Text = Properties.Settings.Default.ImgHeight;
            imgAutoAlignAlignTypeDropDown.SelectedItemIndex = Properties.Settings.Default.imgAutoAlignAlignType;
            excludeTextcheckBox.Checked = Properties.Settings.Default.imgAutoAlighExcludeText;

            // insertMarkdown
            toggleBackgroundCheckBox.Checked = Properties.Settings.Default.ToggleBackground;

            // Add event handlers for text changed events
            fontNameEditBox.TextChanged += SaveSettings;
            fontSizeEditBox.TextChanged += SaveSettings;
            distanceFromBottomEditBox.TextChanged += SaveSettings;
            titleTextEditBox.TextChanged += SaveSettings;
            autoGroupCheckBox.Click += SaveSettings;

            labelOffsetXEditBox.TextChanged += SaveSettings;
            labelOffsetYEditBox.TextChanged += SaveSettings;
            labelTemplateComboBox.TextChanged += SaveSettings;
            labelFontNameEditBox.TextChanged += SaveSettings;
            labelFontSizeEditBox.TextChanged += SaveSettings;

            imgAutoAlignSortTypeDropDown.SelectionChanged += SaveSettings;
            imgAutoAlign_colNum.TextChanged += SaveSettings;
            imgAutoAlign_colSpace.TextChanged += SaveSettings;
            imgAutoAlign_rowSpace.TextChanged += SaveSettings;
            imgWidthEditBpx.TextChanged += SaveSettings;
            imgHeightEditBox.TextChanged += SaveSettings;
            imgAutoAlignAlignTypeDropDown.SelectionChanged += SaveSettings;
            excludeTextcheckBox.Click += SaveSettings;

            toggleBackgroundCheckBox.Click += SaveSettings;
        }

        private void SaveSettings(object sender, RibbonControlEventArgs e)
        {
            // Save Image Title Settings
            Properties.Settings.Default.TitleFontName = fontNameEditBox.Text;
            Properties.Settings.Default.TitleFontSize = fontSizeEditBox.Text;
            Properties.Settings.Default.TitleDistanceFromBottom = distanceFromBottomEditBox.Text;
            Properties.Settings.Default.TitleText = titleTextEditBox.Text;
            Properties.Settings.Default.AutoGroup = autoGroupCheckBox.Checked;

            // Save Image Label Settings
            Properties.Settings.Default.LabelOffsetX = labelOffsetXEditBox.Text;
            Properties.Settings.Default.LabelOffsetY = labelOffsetYEditBox.Text;
            Properties.Settings.Default.LabelTemplate = labelTemplateComboBox.Text;
            Properties.Settings.Default.LabelFontName = labelFontNameEditBox.Text;
            Properties.Settings.Default.LabelFontSize = labelFontSizeEditBox.Text;

            // Save Image Auto Align Settings
            Properties.Settings.Default.imgAutoAlignSortType = imgAutoAlignSortTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.ColNum = imgAutoAlign_colNum.Text;
            Properties.Settings.Default.ColSpace = imgAutoAlign_colSpace.Text;
            Properties.Settings.Default.RowSpace = imgAutoAlign_rowSpace.Text;
            Properties.Settings.Default.ImgWidth = imgWidthEditBpx.Text;
            Properties.Settings.Default.ImgHeight = imgHeightEditBox.Text;
            Properties.Settings.Default.imgAutoAlignAlignType = imgAutoAlignAlignTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.imgAutoAlighExcludeText = excludeTextcheckBox.Checked;

            // Save insertMarkdwon
            Properties.Settings.Default.ToggleBackground = toggleBackgroundCheckBox.Checked;

            // Save all settings
            Properties.Settings.Default.Save();

            // 弹窗显示已保存
            // MessageBox.Show("设置已保存");
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
                        // 不自动换行
                        //titleShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
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
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue; // Lock aspect ratio
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
                    shape.LockAspectRatio = Office.MsoTriState.msoTrue; // Lock aspect ratio
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
                copiedLeft.Clear();
                copiedTop.Clear();
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    copiedLeft.Add(shape.Left + shape.Width / 2);
                    copiedTop.Add(shape.Top + shape.Height / 2);
                }
                // MessageBox.Show($"Copied positions of {sel.ShapeRange.Count} shapes");
            }
            else
            {
                MessageBox.Show("Please select shapes to copy positions.");
            }
        }

        private void pastePosition_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                int count = Math.Min(sel.ShapeRange.Count, copiedLeft.Count);
                if (count == 0)
                {
                    MessageBox.Show("No positions copied yet.");
                    return;
                }

                for (int i = 0; i < count; i++)
                {
                    sel.ShapeRange[i + 1].Left = copiedLeft[i] - sel.ShapeRange[i + 1].Width / 2;
                    sel.ShapeRange[i + 1].Top = copiedTop[i] - sel.ShapeRange[i + 1].Height / 2;
                }

                // if (sel.ShapeRange.Count > copiedLeft.Count)
                // {
                //     MessageBox.Show("More shapes selected than positions copied. Only the first " + copiedLeft.Count + " shapes were positioned.");
                // }
            }
            else
            {
                MessageBox.Show("Please select shapes to paste positions.");
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
                float customWidth = 0;
                float customHeight = 0;

                // Input validation
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
                    rowSpace = colSpace;
                }

                bool useCustomWidth = float.TryParse(imgWidthEditBpx.Text, out customWidth) && customWidth > 0;
                bool useCustomHeight = float.TryParse(imgHeightEditBox.Text, out customHeight) && customHeight > 0;
                var selectedImgShape = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    // Skip text boxes if excludeTextcheckBox is checked
                    if (excludeTextcheckBox.Checked && shape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        continue;
                    }
                    selectedImgShape.Add(shape);
                }

                List<PowerPoint.Shape> shapesToArrange = new List<PowerPoint.Shape>();


                if (imgAutoAlignSortTypeDropDown.SelectedItemIndex == 0)
                {
                    // Create groups based on vertical position
                    var groups = new List<ImageGroup>();
                    var shapes = new List<PowerPoint.Shape>();
                    foreach (PowerPoint.Shape shape in selectedImgShape)
                    {
                        shapes.Add(shape);
                    }

                    // Group shapes based on vertical overlap
                    foreach (var shape in shapes)
                    {
                        bool addedToExistingGroup = false;
                        foreach (var group in groups)
                        {
                            if (group.OverlapsWith(shape))
                            {
                                group.AddShape(shape);
                                addedToExistingGroup = true;
                                break;
                            }
                        }

                        if (!addedToExistingGroup)
                        {
                            var newGroup = new ImageGroup();
                            newGroup.AddShape(shape);
                            groups.Add(newGroup);
                        }
                    }

                    // Sort shapes within each group by x position
                    foreach (var group in groups)
                    {
                        group.Shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
                    }

                    // Sort groups by MinTop
                    groups.Sort((a, b) => a.MinTop.CompareTo(b.MinTop));

                    // Flatten all shapes from all groups into a single list for arrangement
                    foreach (var group in groups)
                    {
                        shapesToArrange.AddRange(group.Shapes);
                    }
                }
                else
                {
                    // Use shapes in their original order
                    foreach (PowerPoint.Shape shape in selectedImgShape)
                    {
                        shapesToArrange.Add(shape);
                    }
                }
                // Now Align image
                float startX = shapesToArrange[0].Left;
                float currentY = shapesToArrange[0].Top;

                if (imgAutoAlignAlignTypeDropDown.SelectedItemIndex == 0)
                {
                    // 最大宽度整齐排列
                    float maxWidth = 0;
                    if (useCustomWidth)
                    {
                        maxWidth = customWidth;
                    }
                    else if (!useCustomWidth && useCustomHeight)
                    {
                        maxWidth = 0;
                        foreach (var shape in shapesToArrange)
                        {
                            float aspectRatio = shape.Width / shape.Height;
                            shape.Height = customHeight;
                            shape.Width = customHeight * aspectRatio;
                            maxWidth = Math.Max(maxWidth, shape.Width);
                        }
                    }
                    else
                    {
                        foreach (var shape in shapesToArrange)
                        {
                            maxWidth = Math.Max(maxWidth, shape.Width);
                        }
                    }

                    float currentX = startX;
                    float rowMaxHeight = 0;
                    int colCount = 0;

                    foreach (var shape in shapesToArrange)
                    {
                        float aspectRatio = shape.Width / shape.Height;
                        if (useCustomWidth && !useCustomHeight)
                        {
                            shape.Width = customWidth;
                            shape.Height = customWidth / aspectRatio;
                        }
                        else if (!useCustomWidth && useCustomHeight)
                        {
                            shape.Height = customHeight;
                            shape.Width = customHeight * aspectRatio;
                        }
                        else if (useCustomWidth && useCustomHeight)
                        {
                            // 取消锁定纵横比
                            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                            shape.Width = customWidth;
                            shape.Height = customHeight;
                        }

                        if (colCount >= colNum)
                        {
                            colCount = 0;
                            currentX = startX;
                            currentY += rowMaxHeight + rowSpace;
                            rowMaxHeight = 0;
                        }

                        shape.Left = currentX;
                        shape.Top = currentY;
                        rowMaxHeight = Math.Max(rowMaxHeight, shape.Height);
                        currentX += maxWidth + colSpace;
                        colCount++;
                    }
                }
                else if (imgAutoAlignAlignTypeDropDown.SelectedItemIndex == 1)
                {
                    // 统一高度排列
                    float referenceHeight = shapesToArrange[0].Height;
                    if (useCustomWidth && !useCustomHeight)
                    {
                        referenceHeight = 0;
                    }
                    float currentX = startX;
                    float rowMaxHeight = 0;
                    int colCount = 0;

                    foreach (var shape in shapesToArrange)
                    {
                        // 保持宽高比调整高度
                        float aspectRatio = shape.Width / shape.Height;
                        if (!useCustomWidth && !useCustomHeight)
                        {

                            shape.Height = referenceHeight;
                            shape.Width = referenceHeight * aspectRatio;
                        }
                        else
                        {
                            if (useCustomWidth && !useCustomHeight)
                            {
                                shape.Width = customWidth;
                                shape.Height = customWidth / aspectRatio;
                            }
                            else if (!useCustomWidth && useCustomHeight)
                            {
                                shape.Height = customHeight;
                                shape.Width = customHeight * aspectRatio;
                                referenceHeight = customHeight;
                            }
                            else
                            {
                                // 取消锁定纵横比
                                shape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                shape.Width = customWidth;
                                shape.Height = customHeight;
                            }
                        }

                        if (colCount >= colNum)
                        {
                            colCount = 0;
                            currentX = startX;
                            currentY += referenceHeight + rowSpace;
                            if (useCustomWidth && !useCustomHeight)
                            {
                                referenceHeight = 0;
                            }
                        }

                        shape.Left = currentX;
                        shape.Top = currentY;
                        currentX += shape.Width + colSpace;
                        colCount++;

                        // Calculate the maximum height in the current row
                        if (useCustomWidth && !useCustomHeight)
                        {
                            referenceHeight = Math.Max(referenceHeight, shape.Height);
                        }
                    }
                }
                else
                {
                    // 瀑布流排列：统一所有图片宽度
                    float[] columnTops = new float[colNum];
                    float[] columnLefts = new float[colNum];

                    // 统一所有图片的宽度
                    float uniformWidth = customWidth > 0 ? customWidth : shapesToArrange[0].Width;

                    // 初始化每列的位置
                    for (int i = 0; i < colNum; i++)
                    {
                        columnTops[i] = currentY;
                        columnLefts[i] = startX + i * (uniformWidth + colSpace);
                    }

                    foreach (var shape in shapesToArrange)
                    {
                        // 统一宽度，保持宽高比
                        float aspectRatio = shape.Width / shape.Height;
                        shape.Width = uniformWidth;
                        shape.Height = uniformWidth / aspectRatio;

                        // 找到高度最小的列
                        int minColumn = 0;
                        float minHeight = columnTops[0];
                        for (int i = 1; i < colNum; i++)
                        {
                            if (columnTops[i] < minHeight)
                            {
                                minHeight = columnTops[i];
                                minColumn = i;
                            }
                        }

                        // 放置图片
                        shape.Left = columnLefts[minColumn];
                        shape.Top = columnTops[minColumn];

                        // 更新该列的高度，加上图片高度和行间距
                        columnTops[minColumn] += shape.Height + rowSpace;
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
                    textBox.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked ?
                        ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                        ColorTranslator.ToOle(Color.White);
                    textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
                    textBox.Line.Weight = 1;

                    // Set the code without language markers
                    textBox.TextFrame.TextRange.Text = code;

                    // Apply base formatting
                    textBox.TextFrame.TextRange.Font.Name = "Consolas";
                    textBox.TextFrame.TextRange.Font.Size = 12;
                    textBox.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked ?
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
                    var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
                    highlighter.ApplyHighlighting(textBox, code, language);

                    // Auto-size the textbox to fit content
                    textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        // Update background color
                        shape.Fill.Solid();
                        shape.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked ?
                            ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                            ColorTranslator.ToOle(Color.White);

                        // Update text color
                        shape.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked ?
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

                        float currentTop = 0;  // Starting position
                        float left = (slide.Master.Width - 500) / 2; // Center horizontally

                        foreach (var segment in segments)
                        {
                            try
                            {
                                PowerPoint.Shape shape = null;
                                if (segment.IsCodeBlock)
                                {
                                    shape = InsertCodeBlock(segment.Content, segment.Language, left, currentTop);
                                }
                                else if (segment.IsTable)
                                {
                                    shape = InsertTable(segment.Content, left, currentTop);
                                }
                                else if (segment.IsMathBlock)
                                {
                                    shape = InsertMathBlock(segment.Content, left, currentTop);
                                }
                                else if (segment.IsBlockQuote)
                                {
                                    shape = InsertBlockQuote(segment.Content, left, currentTop);
                                }
                                else
                                {
                                    string html = ProcessMarkdown(segment.Content);
                                    if (!string.IsNullOrEmpty(html))
                                    {
                                        // Add retry mechanism for clipboard operations
                                        int retryCount = 3;
                                        while (retryCount > 0)
                                        {
                                            try
                                            {
                                                CopyHtmlToClipBoard(segment.Content, html);
                                                System.Threading.Thread.Sleep(100); // Add 100ms delay
                                                PowerPoint.ShapeRange textContent = slide.Shapes.Paste();

                                                if (textContent != null && textContent.Count > 0)
                                                {
                                                    PowerPoint.Shape textShape = textContent[1];
                                                    textShape.Width = 500;
                                                    textShape.Left = left;
                                                    textShape.Top = currentTop;
                                                    currentTop += textShape.Height + 10;

                                                    // Process inline math formulas
                                                    ProcessInlineMathFormulas(textShape);

                                                    if (textShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                                    {
                                                        PowerPoint.TextRange textRange = textShape.TextFrame.TextRange;
                                                        foreach (PowerPoint.TextRange paragraph in textRange.Paragraphs(-1))  // Changed this line
                                                        {
                                                            if (paragraph.ParagraphFormat.Bullet.Type != PowerPoint.PpBulletType.ppBulletNone)
                                                            {
                                                                PowerPoint.PpBulletType ppBulletType = paragraph.ParagraphFormat.Bullet.Type;
                                                                paragraph.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                                                                paragraph.ParagraphFormat.Bullet.Type = ppBulletType;

                                                                // Handle task list items
                                                                string text = paragraph.Text.Trim();
                                                                if (text.StartsWith("- [x]"))
                                                                {
                                                                    char myCharacter = (char)9745; // ☑
                                                                    paragraph.ParagraphFormat.Bullet.Character = myCharacter;
                                                                    paragraph.Text = text.Substring(5).Trim(); // Remove "- [x]"
                                                                }
                                                                else if (text.StartsWith("- [ ]"))
                                                                {
                                                                    char myCharacter = (char)9744; // ☐
                                                                    paragraph.ParagraphFormat.Bullet.Character = myCharacter;
                                                                    paragraph.Text = text.Substring(5).Trim(); // Remove "- [ ]"
                                                                }
                                                            }
                                                        }
                                                    }
                                                    break; // Success, exit retry loop
                                                }
                                            }
                                            catch (System.Runtime.InteropServices.COMException)
                                            {
                                                retryCount--;
                                                if (retryCount <= 0)
                                                {
                                                    MessageBox.Show($"无法粘贴内容: {segment.Content.Substring(0, Math.Min(30, segment.Content.Length))}...");
                                                }
                                                System.Threading.Thread.Sleep(200); // Wait longer before retry
                                            }
                                        }

                                    }
                                }

                                if (shape != null)
                                {
                                    currentTop += shape.Height + 10;
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"处理段落时出错: {ex.Message}");
                                continue; // Continue with next segment
                            }
                        }
                        inputDialog.Dispose();
                        Clipboard.Clear();
                        var dataObject = new DataObject();
                        dataObject.SetData(DataFormats.UnicodeText, markdown);
                        Clipboard.SetDataObject(dataObject, true, 3, 100); // Add retry and timeout parameters
                    }
                }

                inputDialog.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作过程中出错: {ex.Message}\n\n{ex.StackTrace}");
            }
        }

        private void ProcessInlineMathFormulas(PowerPoint.Shape textShape)
        {
            PowerPoint.TextRange textRange = textShape.TextFrame.TextRange;
            string text = textRange.Text;
            // Regex pattern to find math expressions between $ signs 
            var matches = System.Text.RegularExpressions.Regex.Matches(text, @"\$([^$\n]+?)\$");
            // matches.Count如果=0，说明没有匹配到，直接返回
            if (matches.Count == 0)
            {
                return;
            }
            // 创建tempShape，如果不创建，行内数学公式包括分式就不会正常转化
            PowerPoint.Shape tempShape = InsertMathBlock("a", 0, 0);
            // 删除mathShape
            tempShape.Delete();

            // Process matches in reverse order to maintain correct indices
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                int start = match.Index + 1;  // Include the first $ 
                int length = match.Length + 1;  // Include both $ signs
                string formula = match.Groups[1].Value;
                // 替换文本：$公式$为公式
                PowerPoint.TextRange selectedRange = textRange.Characters(start, length);
                selectedRange.Text = formula.Trim('$');
                selectedRange.Select();
                app.CommandBars.ExecuteMso("EquationInsertNew");

                app.CommandBars.ExecuteMso("EquationProfessional");
            }
        }

        private PowerPoint.Shape InsertCodeBlock(string code, string language, float left, float top)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 300);

            // Set code block style
            textBox.Fill.Solid();
            textBox.Fill.ForeColor.RGB = toggleBackgroundCheckBox.Checked ?
                ColorTranslator.ToOle(Color.FromArgb(30, 30, 30)) :
                ColorTranslator.ToOle(Color.White);
            textBox.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            textBox.Line.Weight = 1;

            textBox.TextFrame.TextRange.Text = code;

            // Apply base formatting
            textBox.TextFrame.TextRange.Font.Name = "Consolas";
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.TextFrame.TextRange.Font.Color.RGB = toggleBackgroundCheckBox.Checked ?
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
            var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
            highlighter.ApplyHighlighting(textBox, code, language);

            // Auto-size the textbox to fit content
            textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;

            return textBox;
        }
        public class ImageGroup
        {
            public List<PowerPoint.Shape> Shapes { get; set; } = new List<PowerPoint.Shape>();
            public float MinTop { get; set; }
            public float MaxBottom { get; set; }

            public bool OverlapsWith(PowerPoint.Shape shape)
            {
                float shapeHeight = shape.Height;
                float threshold = shapeHeight * 0.5f; // 50% of shape height
                float shapeBottom = shape.Top + shapeHeight;

                // Calculate overlap height
                float overlapStart = Math.Max(MinTop, shape.Top);
                float overlapEnd = Math.Min(MaxBottom, shapeBottom);
                float overlapHeight = overlapEnd - overlapStart;

                return overlapHeight >= threshold;
            }

            public void AddShape(PowerPoint.Shape shape)
            {
                if (Shapes.Count == 0)
                {
                    MinTop = shape.Top;
                    MaxBottom = shape.Top + shape.Height;
                }
                else
                {
                    MinTop = Math.Min(MinTop, shape.Top);
                    MaxBottom = Math.Max(MaxBottom, shape.Top + shape.Height);
                }
                Shapes.Add(shape);
            }
        }
        private class MarkdownSegment
        {
            public string Content { get; set; }
            public bool IsCodeBlock { get; set; }
            public bool IsTable { get; set; }
            public bool IsMathBlock { get; set; }
            public bool IsBlockQuote { get; set; }  // Add this line
            public string Language { get; set; }
        }

        private List<MarkdownSegment> SplitMarkdownIntoSegments(string markdown)
        {
            var segments = new List<MarkdownSegment>();
            var currentPosition = 0;

            // Updated pattern to better handle tables
            // 1. Tables must start with a header line
            // 2. Followed by a separator line
            // 3. Then one or more data lines
            var pattern = @"(?:```(\w*)\r?\n(.*?)\r?\n```)|" +  // Code blocks
                         @"(?:\|[^\n]*\|\r?\n\|[-|\s]*\|\r?\n(?:\|[^\n]*\|\r?\n)*\|[^\n]*\|?)|" +  // Tables
                         @"(\$\$[\s\S]*?\$\$)|" +               // Math blocks
                        @"(?:(?:^|\n)(?:>[^\n]*(?:\r?\n>[^\n]*)*))";  // 引述块（修改后的模式）

            var regex = new System.Text.RegularExpressions.Regex(pattern,
                System.Text.RegularExpressions.RegexOptions.Multiline |
                System.Text.RegularExpressions.RegexOptions.Singleline);

            var matches = regex.Matches(markdown);

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                // Add text before special block if exists
                if (match.Index > currentPosition)
                {
                    string textBefore = markdown.Substring(currentPosition, match.Index - currentPosition);
                    if (!string.IsNullOrWhiteSpace(textBefore))
                    {
                        segments.Add(new MarkdownSegment
                        {
                            Content = textBefore.Trim(),
                            IsCodeBlock = false,
                            IsTable = false,
                            IsMathBlock = false,
                            IsBlockQuote = false
                        });
                    }
                }

                string content = match.Value;

                // Determine block type and add segment
                if (content.StartsWith("```"))
                {
                    var lines = content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    var language = lines[0].Substring(3).Trim();
                    var codeContent = string.Join("\n", lines.Skip(1).Take(lines.Length - 2));

                    segments.Add(new MarkdownSegment
                    {
                        Content = codeContent,
                        Language = string.IsNullOrEmpty(language) ? "text" : language,
                        IsCodeBlock = true,
                        IsTable = false,
                        IsMathBlock = false,
                        IsBlockQuote = false
                    });
                }
                else if (content.StartsWith("|"))
                {
                    // Clean up table content (remove trailing whitespace and newlines)
                    content = string.Join("\n",
                        content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line.Trim())
                            .Where(line => line.StartsWith("|") && line.EndsWith("|")));

                    segments.Add(new MarkdownSegment
                    {
                        Content = content,
                        IsCodeBlock = false,
                        IsTable = true,
                        IsMathBlock = false,
                        IsBlockQuote = false
                    });
                }
                else if (content.StartsWith("$$"))
                {
                    content = string.Join("\n",
                        content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
                    segments.Add(new MarkdownSegment
                    {
                        Content = content.Replace("\n", ""), // Remove line breaks
                        IsCodeBlock = false,
                        IsTable = false,
                        IsMathBlock = true,
                        IsBlockQuote = false
                    });
                }
                else if (content.TrimStart('\r', '\n').StartsWith(">"))
                {
                    // Clean up block quote content
                    content = string.Join("\n",
                        content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line.TrimStart('>', ' ')));

                    segments.Add(new MarkdownSegment
                    {
                        Content = content,
                        IsCodeBlock = false,
                        IsTable = false,
                        IsMathBlock = false,
                        IsBlockQuote = true
                    });
                }

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
                        IsCodeBlock = false,
                        IsTable = false,
                        IsMathBlock = false,
                        IsBlockQuote = false
                    });
                }
            }

            return segments;
        }

        private PowerPoint.Shape InsertTable(string tableContent, float left, float top)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 300);

            // Convert markdown table to HTML
            // Configure the pipeline with all advanced extensions active
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            string html = Markdown.ToHtml(tableContent, pipeline);
            html = html.Replace("<table>", "<table style='width:500px; border-collapse:collapse;border:1pt solid black;'>");
            html = html.Replace("<td>", "<td style='border:1pt solid black;'>");
            html = html.Replace("<th>", "<th style='border:1pt solid black;'>");

            // Create a temporary DataObject for the table content

            CopyHtmlToClipBoard(tableContent, html);
            System.Threading.Thread.Sleep(100);

            PowerPoint.ShapeRange tableShape = slide.Shapes.Paste();
            if (tableShape != null && tableShape.Count > 0)
            {
                tableShape[1].Left = left;
                tableShape[1].Top = top;
                textBox.Delete();
                return tableShape[1];
            }

            return textBox;
        }

        private PowerPoint.Shape InsertMathBlock(string mathContent, float left, float top)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

            // Insert a new textbox
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 500);

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
            equationShape2.TextFrame.TextRange.Characters(1, equationShape2.TextFrame.TextRange.Text.Length - 1).Text = mathContent;

            // Convert to professional format
            app.CommandBars.ExecuteMso("EquationProfessional");
            // Auto-size and position
            equationShape.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            equationShape.Left = left;
            equationShape.Top = top;



            return equationShape;
        }

        private PowerPoint.Shape InsertBlockQuote(string content, float left, float top)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 300);

            // Configure Markdown pipeline
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();

            // Convert to HTML and remove blockquote tags
            string html = Markdown.ToHtml(content, pipeline)
                .Replace("<blockquote>", "")
                .Replace("</blockquote>", "");

            // Add custom styling
            html = $"<div style='font-family: 微软雅黑; padding: 10px;'>{html}</div>";

            // Copy to clipboard and paste
            CopyHtmlToClipBoard(content, html);
            System.Threading.Thread.Sleep(100);

            PowerPoint.ShapeRange quoteShape = slide.Shapes.Paste();
            if (quoteShape != null && quoteShape.Count > 0)
            {
                PowerPoint.Shape shape = quoteShape[1];
                shape.Left = left;
                shape.Top = top;

                // Add black border
                shape.Line.Visible = Office.MsoTriState.msoTrue;
                shape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Black);
                shape.Line.Weight = 1;

                textBox.Delete();
                return shape;
            }

            return textBox;
        }

        private string ProcessMarkdown(string markdown)
        {
            var codeBlockRegex = new System.Text.RegularExpressions.Regex(
                @"```.*?\r?\n(.*?)\r?\n```",
                System.Text.RegularExpressions.RegexOptions.Singleline
            );

            markdown = codeBlockRegex.Replace(markdown, string.Empty);

            // Convert remaining markdown to HTML
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            string html = Markdown.ToHtml(markdown, pipeline);

            // Add checkbox markers after the checkboxes
            html = html.Replace("<input disabled=\"disabled\" type=\"checkbox\" checked=\"checked\" />", "- [x]");
            html = html.Replace("<input disabled=\"disabled\" type=\"checkbox\" />", "- [ ]");
            // Add table styling
            html = html.Replace("<table>", "<table style='width:500px; border-collapse:collapse;border:1pt solid黑色;'>");
            html = html.Replace("<td>", "<td style='border:1pt solid black;'>");
            html = html.Replace("<th>", "<th style='border:1pt solid black;'>");

            html = html.Replace("<li>", "<li style='margin-left: 10px;'>");
            html = html.Replace("<code>", "<span style='color: #C00000; font-family: Consolas;'>");
            html = html.Replace("</code>", "</span>");

            html = $"<div style='font-family: 微软雅黑;'>{html}</div>";

            // 把<span class="math">\(...\)</span>转换$...$
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<span class=""math"">\\\((.+?)\\\)</span>", m => $"${m.Groups[1].Value}$");
            // 弹窗显示
            //MessageBox.Show($"Markdown转换: {html}");
            return html;
        }

        public void CopyHtmlToClipBoard(string markdown, string html)
        {
            try
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

                int retryCount = 3;
                while (retryCount > 0)
                {
                    try
                    {
                        Clipboard.SetDataObject(dataObject, true, 3, 100); // Add retry and timeout parameters
                        break;
                    }
                    catch (Exception)
                    {
                        retryCount--;
                        if (retryCount <= 0) throw;
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"复制到剪贴板时出错: {ex.Message}");
                throw;
            }
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
        private void openDoc_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.yuque.com/achuan-2/blog/etzcergpmb4rr2sk/");
        }
        private void current_Version(object sender, RibbonControlEventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Version version = assembly.GetName().Version;
            MessageBox.Show($"Version {version}", "Current Version");
        }
        private void aboutDeveloper_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("开发者: Achuan-2\n邮箱: achuan-2@outlook.com\nGithub地址：https://github.com/Achuan-2", "关于开发者");
        }

        private void positionSortCheckBox_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void addLabelsButton_Click(object sender, RibbonControlEventArgs e)
        {
            string fontFamily = labelFontNameEditBox.Text; // 修改为使用新控件
            float fontSize;
            if (!float.TryParse(labelFontSizeEditBox.Text, out fontSize)) // 修改为使用新控件
            {
                MessageBox.Show("请输入有效的字体大小。");
                return;
            }
            float labelOffsetX;
            if (!float.TryParse(labelOffsetXEditBox.Text, out labelOffsetX))
            {
                MessageBox.Show("请输入有效的X偏移量。");
                return;
            }
            float labelOffsetY;
            if (!float.TryParse(labelOffsetYEditBox.Text, out labelOffsetY))
            {
                MessageBox.Show("请输入有效的Y偏移量。");
                return;
            }
            string labelTemplate = labelTemplateComboBox.Text;

            AddLabelsToImages(fontFamily, fontSize, labelOffsetX, labelOffsetY, labelTemplate);
        }

        private void AddLabelsToImages(string fontFamily, float fontSize, float labelOffsetX, float labelOffsetY, string labelTemplate)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
            {
                MessageBox.Show("请选择要添加标签的图片。");
                return;
            }

            var templates = new Dictionary<string, string>
            {
                { "A", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a", "abcdefghijklmnopqrstuvwxyz" },
                { "A)", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" },
                { "a)", "abcdefghijklmnopqrstuvwxyz" },
                { "1", "123456789" },  // Added numeric template
                { "1)", "123456789" }  // Added numeric template with parenthesis
            };

            if (!templates.ContainsKey(labelTemplate))
            {
                labelTemplate = "A";
            }

            string labels = templates[labelTemplate];
            bool isNumeric = labelTemplate.StartsWith("1");
            int selectionCount = sel.ShapeRange.Count;


            // Create groups based on vertical position
            var groups = new List<ImageGroup>();
            var selectedImgShapes = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in sel.ShapeRange)
            {
                // Skip text boxes if excludeTextcheckBox is checked
                if (shape.Type == Office.MsoShapeType.msoTextBox || shape.Type == Office.MsoShapeType.msoAutoShape)
                {
                    continue;
                }
                selectedImgShapes.Add(shape);
            }
            if (selectedImgShapes.Count == 0)
            {
                MessageBox.Show("请选择要添加标签的图片。");
                return;
            }

            // Group shapes based on vertical overlap
            foreach (var shape in selectedImgShapes)
            {
                bool addedToExistingGroup = false;
                foreach (var group in groups)
                {
                    if (group.OverlapsWith(shape))
                    {
                        group.AddShape(shape);
                        addedToExistingGroup = true;
                        break;
                    }
                }

                if (!addedToExistingGroup)
                {
                    var newGroup = new ImageGroup();
                    newGroup.AddShape(shape);
                    groups.Add(newGroup);
                }
            }

            // Sort shapes within each group by x position
            foreach (var group in groups)
            {
                group.Shapes.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            // Sort groups by MinTop
            groups.Sort((a, b) => a.MinTop.CompareTo(b.MinTop));

            // Create flattened list of sorted shapes
            var sortedShapes = new List<PowerPoint.Shape>();
            foreach (var group in groups)
            {
                sortedShapes.AddRange(group.Shapes);
            }

            // Add labels to sorted shapes
            for (int i = 0; i < sortedShapes.Count; i++)
            {
                try
                {
                    var item = sortedShapes[i];
                    string label;
                    if (isNumeric)
                    {
                        label = (i + 1).ToString();
                    }
                    else
                    {
                        label = labels[i % labels.Length].ToString();
                    }

                    if (labelTemplate.EndsWith(")"))
                    {
                        label += ")";
                    }

                    var textBox = app.ActiveWindow.View.Slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        item.Left + labelOffsetX,
                        item.Top + labelOffsetY,
                        0, // Initial width
                        fontSize * 2);

                    // Set the text and font properties
                    textBox.TextFrame.TextRange.Text = label;
                    textBox.TextFrame.TextRange.Font.Size = fontSize;
                    textBox.TextFrame.TextRange.Font.Name = fontFamily;
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;

                    // Auto-size the textbox to fit the text
                    textBox.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    // 不自动换行
                    textBox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"添加标签时出错: {ex.Message}");
                }
            }
        }
        private static Microsoft.Office.Interop.PowerPoint.Font _copiedFont;


        private void copyStyle_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sourceShape = sel.ShapeRange[1];
                // 捕获格式
                sourceShape.PickUp();
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                _copiedFont = sel.TextRange.Font;
            }
            else
            {
            }

        }

        private void pasteStyle_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    shape.Apply();
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                ApplyFont(sel.TextRange.Font);
            }
            else
            {
            }
        }

        private void ApplyFont(Microsoft.Office.Interop.PowerPoint.Font targetFont)
        {
            targetFont.Name = _copiedFont.Name;
            targetFont.Size = _copiedFont.Size;
            targetFont.Bold = _copiedFont.Bold;
            targetFont.Italic = _copiedFont.Italic;
            targetFont.Color.RGB = _copiedFont.Color.RGB;
            targetFont.Underline = _copiedFont.Underline;
        }
    }


}

