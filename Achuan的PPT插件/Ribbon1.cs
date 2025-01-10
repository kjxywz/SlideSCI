using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace Achuan的PPT插件
{
    public partial class Ribbon1
    {
        PowerPoint.Application app;
        private float copiedWidth;
        private float copiedHeight;
        private float copiedLeft;
        private float copiedTop;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            app = Globals.ThisAddIn.Application;
        }

        private void AddTitleToImage(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    // Create textbox below the image
                    PowerPoint.Shape textbox = app.ActiveWindow.View.Slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        shape.Left,
                        shape.Top + shape.Height,
                        shape.Width,
                        20);

                    // Set text properties
                    textbox.TextFrame.TextRange.Font.Name = "微软雅黑";
                    textbox.TextFrame.TextRange.Font.Size = 14;
                    textbox.TextFrame.TextRange.Text = "图片标题";
                    
                    // Center align the text
                    textbox.TextFrame.TextRange.ParagraphFormat.Alignment = 
                        PowerPoint.PpParagraphAlignment.ppAlignCenter;

                    // Group the image and the title
                    PowerPoint.ShapeRange shapeRange = app.ActiveWindow.Selection.SlideRange.Shapes.Range(new string[] { shape.Name, textbox.Name });
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
                if (!useCustomWidth)
                {
                    imgWidth = firstShape.Width;
                }

                float startX = firstShape.Left;
                float startY = firstShape.Top;
                float currentX = startX;
                float currentY = startY;
                int currentCol = 0;

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (useCustomWidth)
                    {
                        shape.Width = imgWidth;
                    }
                    if (useCustomHeight)
                    {
                        shape.Height = imgHeight;
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
    }
}