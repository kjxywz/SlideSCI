using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Drawing;

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

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void AddTitleToImage(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Shape selectedShape = app.ActiveWindow.Selection.ShapeRange[1];

            if (selectedShape != null && selectedShape.Type == Office.MsoShapeType.msoPicture)
            {
                float fontSize = float.Parse(fontSizeEditBox.Text);
                float distanceFromBottom = float.Parse(distanceFromBottomEditBox.Text);
                bool autoGroup = autoGroupCheckBox.Checked;
                string fontName = fontNameEditBox.Text;
                string titleText = titleTextEditBox.Text;

                PowerPoint.Shape titleShape = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    selectedShape.Left,
                    selectedShape.Top + selectedShape.Height + distanceFromBottom,
                    selectedShape.Width,
                    fontSize * 2);

                titleShape.TextFrame.TextRange.Text = titleText;
                titleShape.TextFrame.TextRange.Font.Size = fontSize;
                titleShape.TextFrame.TextRange.Font.NameFarEast = fontName;; // Ensure FarEast font is set
                titleShape.TextFrame.TextRange.Font.Name = fontName; // Ensure font is set
                titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                if (autoGroup)
                {
                    PowerPoint.ShapeRange shapeRange = slide.Shapes.Range(new string[] { selectedShape.Name, titleShape.Name });
                    shapeRange.Group();
                }
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
                Text = "插入代码块"
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
    }
}