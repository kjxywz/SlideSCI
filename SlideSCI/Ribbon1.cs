using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Markdig;  // 修改为使用签名版本的命名空间
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Font = System.Drawing.Font; // Add this line
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private PowerPoint.Application app;
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
            labelBoldcheckBox.Checked = Properties.Settings.Default.LabelBold;

            // Load Image Auto Align Settings
            imgAutoAlignSortTypeDropDown.SelectedItemIndex = Properties.Settings.Default.imgAutoAlignSortType;
            imgAutoAlign_colNum.Text = Properties.Settings.Default.ColNum;
            imgAutoAlign_colSpace.Text = Properties.Settings.Default.ColSpace;
            imgAutoAlign_rowSpace.Text = Properties.Settings.Default.RowSpace;
            imgWidthEditBpx.Text = Properties.Settings.Default.ImgWidth;
            imgHeightEditBox.Text = Properties.Settings.Default.ImgHeight;
            imgAutoAlignAlignTypeDropDown.SelectedItemIndex = Properties.Settings.Default.imgAutoAlignAlignType;
            excludeTextcheckBox.Checked = Properties.Settings.Default.imgAutoAlighExcludeText;
            excludeTextcheckBox2.Checked = Properties.Settings.Default.imgAddTitleExcludeText;

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

            labelBoldcheckBox.Click += SaveSettings;

            toggleBackgroundCheckBox.Click += SaveSettings;
            iniCombobox();


        }

        /// <summary>
        /// 初始化下拉框的值
        /// </summary>
        public void iniCombobox()
        {
            //字体名
            List<string> FontNames = new List<string>() { "Arial", "微软雅黑", "黑体", "方正兰亭黑体", "仿宋", "楷体", "宋体", "新宋体", "华文中宋", "华文仿宋", "华文行楷", "华文新魏", "汉仪综艺体简", "思源黑体", "思源宋体", "庞门正道标题体", "方正清刻本悦宋", "文悦新青年体", "演示新手书" };
            FreshCombobox(fontNameEditBox, FontNames);
            FreshCombobox(labelFontNameEditBox, FontNames);
            //字号
            List<string> FontSizes = new List<string>() { "2", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "20", "22", "24", "26", "28", "30", "40", "50", "60", "80", "100", "120", "150", "200" };
            FreshCombobox(fontSizeEditBox, FontSizes);
            FreshCombobox(labelFontSizeEditBox, FontSizes);
            //图片宽度和高度
            List<string> PicSizes = new List<string>() { "0cm", "0.5cm", "1cm", "2cm", "3cm", "4cm", "5cm", "6cm", "7cm", "8cm", "9cm", "10cm", "12cm", "15cm", "20cm", "25cm", "30cm", "35cm", "40cm", "45cm", "50cm", "60cm", "70cm", "80cm", "100cm", "120cm", "150cm", "200cm" };
            FreshCombobox(imgWidthEditBpx, PicSizes);
            FreshCombobox(imgHeightEditBox, PicSizes);
            //图下距离
            List<string> PicDistance = new List<string>() { "0", "1", "2", "3", "4", "5", "6", "7", "8", "10", "11", "12", "13", "14", "15", "20", "25", "30", "35", "40", "45", "50", "55", "60", "65", "70", "75", "80", "90", "100", "120", "150", "200", "500" };
            FreshCombobox(distanceFromBottomEditBox, PicDistance);
            //XY偏移
            List<string> OffsetValues = new List<string>() { "-40", "-30", "-20", "-10", "-15", "-10", "-9", "-8", "-7", "-6", "-5", "-4", "-3", "-2", "-1", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "15", "20", "25", "30", "40", };
            FreshCombobox(labelOffsetYEditBox, OffsetValues);
            FreshCombobox(labelOffsetXEditBox, OffsetValues);
            //列间距
            List<string> columnGap = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "30", "35", "40", "45", "50", "55", "60", "80" };
            FreshCombobox(imgAutoAlign_colSpace, columnGap);
            //行间距
            List<string> RowGap = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17(≈08字框高)", "18(≈09字框高)", "20(≈10字框高)", "22(≈12字框高)", "25(≈14字框高)", "27(≈16字框高)", "29(≈18字框高)", "32(≈20字框高)", "34(≈22字框高)", "37(≈24字框高)", "39(≈26字框高)", "41(≈28字框高)", "44(≈30字框高)", "51(≈35字框高)", "56(≈40字框高)", "68(≈50字框高)", "80(≈60字框高)" };
            FreshCombobox(imgAutoAlign_rowSpace, RowGap);
            //列数量
            List<string> columNums = new List<string>() { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20" };
            FreshCombobox(imgAutoAlign_colNum, columNums);

        }

        /// <summary>
        /// RibbonComboBox下拉值初始化
        /// </summary>
        /// <param name="BOX"></param>
        private void FreshCombobox(RibbonComboBox BOX, List<string> itemLabel)
        {
            BOX.Items.Clear();
            // 使用 LINQ 创建 RibbonDropDownItem 并添加到 RibbonDropDown 中
            itemLabel.Select(x =>
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = x; // 设置项的显示文本
                return item;
            }).ToList().ForEach(item => BOX.Items.Add(item)); // 将项添加到下拉菜单中

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
            Properties.Settings.Default.LabelBold = labelBoldcheckBox.Checked;
            // Save Image Auto Align Settings
            Properties.Settings.Default.imgAutoAlignSortType = imgAutoAlignSortTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.ColNum = imgAutoAlign_colNum.Text;
            Properties.Settings.Default.ColSpace = imgAutoAlign_colSpace.Text;
            Properties.Settings.Default.RowSpace = imgAutoAlign_rowSpace.Text;
            Properties.Settings.Default.ImgWidth = imgWidthEditBpx.Text;
            Properties.Settings.Default.ImgHeight = imgHeightEditBox.Text;
            Properties.Settings.Default.imgAutoAlignAlignType = imgAutoAlignAlignTypeDropDown.SelectedItemIndex;
            Properties.Settings.Default.imgAutoAlighExcludeText = excludeTextcheckBox.Checked;
            Properties.Settings.Default.imgAddTitleExcludeText = excludeTextcheckBox2.Checked;

            // Save insertMarkdwon
            Properties.Settings.Default.ToggleBackground = toggleBackgroundCheckBox.Checked;

            // Save all settings
            Properties.Settings.Default.Save();

            // 弹窗显示已保存
            // MessageBox.Show("设置已保存");
        }

        /// <summary>
        /// 图片加标题
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddTitleToImage(object sender, RibbonControlEventArgs e)
        {
            AddTitleFun();
        }

        /// <summary>
        /// 图片加标题
        /// </summary>
        private void AddTitleFun()
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            Slide slide = app.ActiveWindow.View.Slide;
            Selection sel = app.ActiveWindow.Selection;
            bool autoGroup = autoGroupCheckBox.Checked; // 自动编组
            List<ShapeRange> allshapesName = new List<ShapeRange>(); // 需要编组的对象集合
            List<Shape> allshapes = new List<Shape>(); // 编组后的对象

            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
            float fontSize = float.Parse(fontSizeEditBox.Text); // 字号
            float distanceFromBottom = float.Parse(distanceFromBottomEditBox.Text); // 图下距离
            string fontName = fontNameEditBox.Text; // 字体名称
            string titleText = titleTextEditBox.Text; // 标题文本
            int count = 1;
            float tolerance = 10f; // 通常图片排列错位容差，10就够用
            ShapeRange sel2 = GetSortedSelection(sel, tolerance);
            var selectedImgShape = new List<Shape>();

            foreach (Shape shape in sel.ShapeRange)
            {
                Office.MsoShapeType objType = shape.Type;
                // 是否排除文本框、形状等格式。excludeTextcheckBox2.Checked，则排除
                if (excludeTextcheckBox2.Checked && (objType is Office.MsoShapeType.msoTextBox || objType is Office.MsoShapeType.msoAutoShape || objType is Office.MsoShapeType.msoMedia))
                {
                continue;
                }
                selectedImgShape.Add(shape);
            }

            foreach (Shape selectedShape in selectedImgShape)
            {
                Shape titleShape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, selectedShape.Left, selectedShape.Top + selectedShape.Height + distanceFromBottom, selectedShape.Width, fontSize * 2);
                
                // 设置标题文本和格式
                titleShape.TextFrame.TextRange.Text = titleText;
                titleShape.TextFrame.TextRange.Font.Size = fontSize;
                titleShape.TextFrame.TextRange.Font.NameFarEast = fontName; // Ensure FarEast font is set
                titleShape.TextFrame.TextRange.Font.Name = fontName; // Ensure font is set
                titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;


                // 形状中的文字是否自动换行
                titleShape.TextFrame.WordWrap = Office.MsoTriState.msoTrue; 
                // 自动调整文本框大小
                titleShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                // https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe.autosize: 
                // ppAutoSizeNone ：不自动调整大小，ppAutoSizeShapeToFitText ：自动调整大小，ppAutoSizeMixed：混合模式



                // 设置文本框宽度
                titleShape.Width = selectedShape.Width;
                titleShape.Left = selectedShape.Left; // 设置文本框左对齐
                //titleShape.Top = selectedShape.Top + selectedShape.Height + distanceFromBottom;
                

                
                allshapesName.Add(slide.Shapes.Range(new string[] { selectedShape.Name, titleShape.Name }));
                
                // 自动选择
                if (count == 1)
                {
                titleShape.Select(Office.MsoTriState.msoTrue);
                }
                else
                {
                titleShape.Select(Office.MsoTriState.msoFalse);
                }
                count++;
            }
            }
            else
            {
            MessageBox.Show("请选择需要增加标题的图片、形状、视频对象.");
            }

            // 自动编组
            if (autoGroup)
            {
            foreach (var shapeRange2 in allshapesName)
            {
                Shape GroupObj;
                try
                {
                GroupObj = shapeRange2.Group();
                allshapes.Add(GroupObj);
                SelectMultipleShapes(allshapes);
                }
                catch (Exception ex)
                {
                try
                {
                    // 如果编组失败，尝试复制、删除、粘贴再编组
                    // float originalLeft = shapeRange2[1].Left;
                    // float originalTop = shapeRange2[1].Top;

                    shapeRange2.Copy();
                    shapeRange2.Delete();

                    ShapeRange pastedShapes = slide.Shapes.Paste();
                    // pastedShapes.Left = originalLeft;
                    // pastedShapes.Top = originalTop;

                    GroupObj = pastedShapes.Group();
                    allshapes.Add(GroupObj);
                    SelectMultipleShapes(allshapes);
                }
                catch (Exception innerEx)
                {
                    MessageBox.Show($"编组失败：{innerEx.Message}");
                    continue;
                }
                }
            }
            }
        }

        /// <summary>
        /// 选择集排序
        /// </summary>
        /// <param name="initialSelection">原始选择集</param>
        /// <returns></returns>
        public ShapeRange GetSortedSelection(Selection initialSelection, float tolerance)
        {
            try
            {
                // 确保选择集中有形状对象
                if (initialSelection.ShapeRange.Count == 0)
                {
                    MessageBox.Show("初始选择集中未包含任何形状。");
                    return null;
                }

                // 将选择集中的形状转换为 List<PowerPoint.Shape>
                List<Shape> shapes = initialSelection.ShapeRange.Cast<Shape>().ToList();

                // 根据 X 从小到大、Y 从大到小排序
                var sortedShapes = shapes
                    .OrderBy(shape => shape.Top + tolerance)  // Y 坐标从小到大
                    .ThenByDescending(shape => (shape.Left + tolerance) * -1) // X 坐标从大到小
                    .ToList();

                // 将排序后的形状转换为 ShapeRange
                object[] shapeNames = sortedShapes.Select(shape => (object)shape.Name).ToArray();
                ShapeRange sortedShapeRange = initialSelection.ShapeRange[1].Parent.Shapes.Range(shapeNames);

                return sortedShapeRange;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"排序失败：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 选择多个对象
        /// </summary>
        /// <param name="shapeNames"></param>
        public void SelectMultipleShapes(List<Shape> shapesToSelect)
        {
            try
            {
                // 获取当前 PowerPoint 应用实例
                PowerPoint.Application pptApp = Globals.ThisAddIn.Application;

                // 检查是否处于普通视图
                if (pptApp.ActiveWindow.View.Type != PpViewType.ppViewNormal)
                {
                    MessageBox.Show("请切换到普通视图以操作形状。");
                    return;
                }

                // 获取当前幻灯片
                Slide currentSlide = pptApp.ActiveWindow.View.Slide as Slide;
                if (currentSlide == null)
                {
                    MessageBox.Show("未找到活动幻灯片。");
                    return;
                }

                // 提取形状名称列表
                List<object> shapeNames = new List<object>();
                foreach (Shape shape in shapesToSelect)
                {
                    shapeNames.Add((object)shape.Name);
                }

                // 选中所有形状
                if (shapeNames.Count > 0)
                {
                    ShapeRange selectedShapes = currentSlide.Shapes.Range(shapeNames.ToArray());
                    selectedShapes.Select();
                    pptApp.ActiveWindow.Activate(); // 确保窗口焦点
                }
                else
                {
                    MessageBox.Show("未提供有效的形状列表。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败：{ex.Message}");
            }
        }

        private void pasteImgWidthHeight_Click(object sender, RibbonControlEventArgs e)
        {
            if (copiedWidth <= 0 || copiedHeight <= 0)
            {
                MessageBox.Show("Invalid copied dimensions. Please copy the dimensions again.");
                return;
            }

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
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

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];
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

            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                copiedLeft.Clear();
                copiedTop.Clear();
                foreach (Shape shape in sel.ShapeRange)
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
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

        /// <summary>
        /// 图片排列
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void imgAutoAlign_Click(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        /// <summary>
        /// 图片对齐排列
        /// </summary>
        private void AlignPics()
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
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

                if (!float.TryParse(imgAutoAlign_rowSpace.Text.Split(new char[] { '(', ' ' })[0], out rowSpace) || rowSpace < 0)
                {
                    rowSpace = colSpace;
                }

                bool useCustomWidth = float.TryParse(imgWidthEditBpx.Text.Split('c')[0], out customWidth) && customWidth > 0;
                bool useCustomHeight = float.TryParse(imgHeightEditBox.Text.Split('c')[0], out customHeight) && customHeight > 0;
                customWidth = (float)(customWidth * 28.34646);
                customHeight = (float)(customHeight * 28.34646);
                var selectedImgShape = new List<Shape>();
                foreach (Shape shape in sel.ShapeRange)
                {
                    // Skip text boxes if excludeTextcheckBox is checked
                    Office.MsoShapeType objType = shape.Type;
                    if (excludeTextcheckBox.Checked && (objType is Office.MsoShapeType.msoTextBox || objType is Office.MsoShapeType.msoAutoShape || objType is Office.MsoShapeType.msoMedia))
                    {
                        continue;
                    }
                    selectedImgShape.Add(shape);
                }

                List<Shape> shapesToArrange = new List<Shape>();

                if (imgAutoAlignSortTypeDropDown.SelectedItemIndex == 0)
                {
                    // Create groups based on vertical position
                    var groups = new List<ImageGroup>();
                    var shapes = new List<Shape>();
                    foreach (Shape shape in selectedImgShape)
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
                    foreach (Shape shape in selectedImgShape)
                    {
                        shapesToArrange.Add(shape);
                    }
                }
                // Now Align image
                float startX = shapesToArrange[0].Left;
                float startY = shapesToArrange[0].Top;
                float currentY = shapesToArrange[0].Top;

                if (imgAutoAlignAlignTypeDropDown.SelectedItemIndex == 0)
                {
                    // 1. 预先将图片分配到列
                    List<List<Shape>> columns = new List<List<Shape>>();
                    for (int i = 0; i < colNum; i++)
                    {
                        columns.Add(new List<Shape>());
                    }

                    for (int i = 0; i < shapesToArrange.Count; i++)
                    {
                        columns[i % colNum].Add(shapesToArrange[i]); // 按顺序分配到列
                    }

                    // 2. 计算每列的最大宽度
                    List<float> columnWidths = new List<float>();
                    for (int i = 0; i < colNum; i++)
                    {
                        float columnMaxWidth = 0;
                        foreach (var shape in columns[i])
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
                                // 取消锁定纵横比 (假设 Shape 类有 LockAspectRatio 属性)
                                // shape.LockAspectRatio = Office.MsoTriState.msoFalse; // 如果使用 Office Interop
                                shape.Width = customWidth;
                                shape.Height = customHeight;
                            }
                            columnMaxWidth = Math.Max(columnMaxWidth, shape.Width);
                        }
                        columnWidths.Add(columnMaxWidth);
                    }
                    float currentX = startX;
                    float rowMaxHeight = 0;
                    int colCount = 0;
                    // 3. 按行进行排列
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
                            // referenceHeight = customHeight;
                            // 需要计算最大占位宽度
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
                        currentX += columnWidths[colCount] + colSpace;
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
                    Slide slide = app.ActiveWindow.View.Slide;

                    Shape textBox = slide.Shapes.AddTextbox(
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
                        PpParagraphAlignment.ppAlignLeft;

                    // Set margins
                    textBox.TextFrame.MarginLeft = 10;
                    textBox.TextFrame.MarginRight = 10;
                    textBox.TextFrame.MarginTop = 5;
                    textBox.TextFrame.MarginBottom = 5;

                    // Apply syntax highlighting
                    var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
                    highlighter.ApplyHighlighting(textBox, code, language);

                    // Auto-size the textbox to fit content
                    textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                }
            }
        }

        private void checkBox1_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
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
            Slide slide = app.ActiveWindow.View.Slide;

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
                        Shape textBox = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        slide.Master.Width / 2 - 100, slide.Master.Height / 2 - 50, 500, 500);

                        // Select the newly inserted textbox
                        textBox.Select();
                        app.ActiveWindow.Selection.TextRange.Select();

                        // Run SwitchLatex
                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        Shape equationShape = app.ActiveWindow.Selection.ShapeRange[1];
                        equationShape.TextFrame.TextRange.Characters(1, equationShape.TextFrame.TextRange.Text.Length - 1).Text = "\u24C9";

                        app.CommandBars.ExecuteMso("EquationInsertNew");
                        app.ActiveWindow.Selection.TextRange.Select();
                        Shape equationShape2 = app.ActiveWindow.Selection.ShapeRange[1];
                        // Set the LaTeX input to the equation shape
                        equationShape2.TextFrame.TextRange.Characters(1, equationShape2.TextFrame.TextRange.Text.Length - 1).Text = latexInput;

                        // Convert to professional format
                        app.CommandBars.ExecuteMso("EquationProfessional");

                        textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
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
                        Slide slide = app.ActiveWindow.View.Slide;

                        // Split markdown into segments
                        var segments = SplitMarkdownIntoSegments(markdown);

                        float currentTop = slide.Master.Height/2;  // Starting position
                        float left = (slide.Master.Width - 500) / 2; // Center horizontally

                        foreach (var segment in segments)
                        {
                            try
                            {
                                Shape shape = null;
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
                                                ShapeRange textContent = slide.Shapes.Paste();

                                                if (textContent != null && textContent.Count > 0)
                                                {
                                                    Shape textShape = textContent[1];
                                                    textShape.Width = 500;
                                                    textShape.Left = left;
                                                    textShape.Top = currentTop;
                                                    currentTop += textShape.Height + 10;

                                                    // Process inline math formulas
                                                    ProcessInlineMathFormulas(textShape);

                                                    if (textShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                                    {
                                                        TextRange textRange = textShape.TextFrame.TextRange;
                                                        foreach (TextRange paragraph in textRange.Paragraphs(-1))  // Changed this line
                                                        {
                                                            if (paragraph.ParagraphFormat.Bullet.Type != PpBulletType.ppBulletNone)
                                                            {
                                                                // 保存type和start value
                                                                PpBulletType ppBulletType = paragraph.ParagraphFormat.Bullet.Type;
                                                                // 如果是ppBulletNumbered，保存start value
                                                                

                                                                // 获取character
                                                                int character = paragraph.ParagraphFormat.Bullet.Character;
                                                                int startValue = paragraph.ParagraphFormat.Bullet.StartValue; // 有序列表的编号
                                                                PpNumberedBulletStyle stype = paragraph.ParagraphFormat.Bullet.Style; // 有序列表的样式：1、A、一等

                                                                //paragraph.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNone;
                                                                paragraph.ParagraphFormat.Bullet.Type = ppBulletType;
                                                                paragraph.ParagraphFormat.Bullet.Character = character;
                                                                if (ppBulletType == PpBulletType.ppBulletNumbered)
                                                                {
                                                                    paragraph.ParagraphFormat.Bullet.StartValue = startValue;
                                                                    paragraph.ParagraphFormat.Bullet.Style = stype;
                                                                }
                                                                 // 弹窗输出ppBulletType是什么
                                                                // MessageBox.Show($"Bullet type: {ppBulletType}, Start value: {startValue}, Character: {character}, Style: {stype}");
                                                                
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

        private void ProcessInlineMathFormulas(Shape textShape)
        {
            TextRange textRange = textShape.TextFrame.TextRange;
            string text = textRange.Text;
            // Regex pattern to find math expressions between $ signs
            var matches = Regex.Matches(text, @"\$([^$\n]+?)\$");
            // matches.Count如果=0，说明没有匹配到，直接返回
            if (matches.Count == 0)
            {
                return;
            }
            // 创建tempShape，如果不创建，行内数学公式包括分式就不会正常转化
            Shape tempShape = InsertMathBlock("a", 0, 0);
            // 删除mathShape
            tempShape.Delete();

            // Process matches in reverse order to maintain correct indices
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                int start = match.Index + 1;  // 1-based start index of the match (e.g., the first '$')
                int length = match.Length;    // Length of the matched string (e.g., "$formula$")
                string formula = match.Groups[1].Value; // Content within $...$ (e.g., "formula")
                
                // Select the range "$formula$"
                TextRange selectedRange = textRange.Characters(start, length);
                // Replace its text with "formula"
                selectedRange.Text = formula; 
                selectedRange.Select();
                app.CommandBars.ExecuteMso("EquationInsertNew");

                app.CommandBars.ExecuteMso("EquationProfessional");
            }
        }

        private Shape InsertCodeBlock(string code, string language, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
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
                PpParagraphAlignment.ppAlignLeft;

            // Set margins
            textBox.TextFrame.MarginLeft = 10;
            textBox.TextFrame.MarginRight = 10;
            textBox.TextFrame.MarginTop = 5;
            textBox.TextFrame.MarginBottom = 5;

            // Apply syntax highlighting
            var highlighter = new CodeHighlighter(toggleBackgroundCheckBox.Checked);
            highlighter.ApplyHighlighting(textBox, code, language);

            // Auto-size the textbox to fit content
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;

            return textBox;
        }

        public class ImageGroup
        {
            public List<Shape> Shapes { get; set; } = new List<Shape>();
            public float MinTop { get; set; }
            public float MaxBottom { get; set; }

            public bool OverlapsWith(Shape shape)
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

            public void AddShape(Shape shape)
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

            var regex = new Regex(pattern,
                RegexOptions.Multiline |
                RegexOptions.Singleline);

            var matches = regex.Matches(markdown);

            foreach (Match match in matches)
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

        private Shape InsertTable(string tableContent, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
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

            ShapeRange tableShape = slide.Shapes.Paste();
            if (tableShape != null && tableShape.Count > 0)
            {
                tableShape[1].Left = left;
                tableShape[1].Top = top;
                textBox.Delete();
                return tableShape[1];
            }

            return textBox;
        }

        private Shape InsertMathBlock(string mathContent, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;

            // Insert a new textbox
            Shape textBox = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left, top, 500, 500);

            // Select the newly inserted textbox
            textBox.Select();
            app.ActiveWindow.Selection.TextRange.Select();

            // Run SwitchLatex
            app.CommandBars.ExecuteMso("EquationInsertNew");
            Shape equationShape = app.ActiveWindow.Selection.ShapeRange[1];
            equationShape.TextFrame.TextRange.Characters(1, equationShape.TextFrame.TextRange.Text.Length - 1).Text = "\u24C9";

            app.CommandBars.ExecuteMso("EquationInsertNew");
            app.ActiveWindow.Selection.TextRange.Select();
            Shape equationShape2 = app.ActiveWindow.Selection.ShapeRange[1];
            // Set the LaTeX input to the equation shape
            equationShape2.TextFrame.TextRange.Characters(1, equationShape2.TextFrame.TextRange.Text.Length - 1).Text = mathContent;

            // Convert to professional format
            app.CommandBars.ExecuteMso("EquationProfessional");
            // Auto-size and position
            equationShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            equationShape.Left = left;
            equationShape.Top = top;

            return equationShape;
        }

        private Shape InsertBlockQuote(string content, float left, float top)
        {
            Slide slide = app.ActiveWindow.View.Slide;
            Shape textBox = slide.Shapes.AddTextbox(
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

            ShapeRange quoteShape = slide.Shapes.Paste();
            if (quoteShape != null && quoteShape.Count > 0)
            {
                Shape shape = quoteShape[1];
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
            var codeBlockRegex = new Regex(
                @"```.*?\r?\n(.*?)\r?\n```",
                RegexOptions.Singleline
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
            html = Regex.Replace(html, @"<span class=""math"">\\\((.+?)\\\)</span>", m => $"${m.Groups[1].Value}$");
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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape shape = sel.ShapeRange[1];

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
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
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
            Version version = assembly.GetName().Version;
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

        /// <summary>
        /// 图片添加标签
        /// </summary>
        /// <param name="fontFamily"></param>
        /// <param name="fontSize"></param>
        /// <param name="labelOffsetX"></param>
        /// <param name="labelOffsetY"></param>
        /// <param name="labelTemplate">标签格式</param>
        private void AddLabelsToImages(string fontFamily, float fontSize, float labelOffsetX, float labelOffsetY, string labelTemplate)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
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
                { "1)", "123456789" },  // Added numeric template with parenthesis
                {"Ⅰ","ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                {"Ⅰ)","ⅠⅡⅢⅣⅤⅥⅦⅦⅨⅩ" },
                {"①","①②③④⑤⑥⑦⑧⑨⑩" },
                {"①)","①②③④⑤⑥⑦⑧⑨⑩" },
                {"一","一二三四五六七八九十" },
                {"一)","一二三四五六七八九十" }

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
            var selectedImgShapes = new List<Shape>();
            foreach (Shape shape in sel.ShapeRange)
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
            var sortedShapes = new List<Shape>();
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
                    textBox.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                    // Auto-size the textbox to fit the text
                    textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    // 不自动换行
                    textBox.TextFrame.WordWrap = Office.MsoTriState.msoFalse;

                    // 自动加粗
                    if (labelBoldcheckBox.Checked)
                    {
                        textBox.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    }
                    // 自动选择
                    if (i == 0)
                    {
                        textBox.Select(Office.MsoTriState.msoTrue);
                    }
                    else
                    {
                        textBox.Select(Office.MsoTriState.msoFalse);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"添加标签时出错: {ex.Message}");
                }
            }
        }

        private static PowerPoint.Font _copiedFont;

        private void copyStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape sourceShape = sel.ShapeRange[1];
                // 捕获格式
                sourceShape.PickUp();
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                _copiedFont = sel.TextRange.Font;
            }
            else
            {
            }
        }

        private void pasteStyle_Click(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (Shape shape in sel.ShapeRange)
                {
                    shape.Apply();
                }
            }
            else if (sel.Type == PpSelectionType.ppSelectionText)
            {
                ApplyFont(sel.TextRange.Font);
            }
            else
            {
            }
        }

        private void ApplyFont(PowerPoint.Font targetFont)
        {
            targetFont.Name = _copiedFont.Name;
            targetFont.Size = _copiedFont.Size;
            targetFont.Bold = _copiedFont.Bold;
            targetFont.Italic = _copiedFont.Italic;
            targetFont.Color.RGB = _copiedFont.Color.RGB;
            targetFont.Underline = _copiedFont.Underline;
        }

        private void pastePictureAndText(object sender, RibbonControlEventArgs e)
        {
            Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                try
                {
                    // Store original position
                    // float left = sel.ShapeRange.Left;
                    // float top = sel.ShapeRange.Top;

                    // Group the shapes first if multiple shapes selected
                    Shape groupedShape;
                    try {
                        // First attempt - try to group directly
                        groupedShape = sel.ShapeRange.Count > 1 ? 
                            sel.ShapeRange.Group() : sel.ShapeRange[1];
                    } 
                    catch (Exception ex) {
                        // If direct grouping fails, try the copy-delete-paste-group approach
                        try {
                            // Copy the shapes
                            sel.ShapeRange.Copy();
                            
   
                            
                            // Delete original shapes
                            sel.ShapeRange.Delete();
                            
                            // Paste back the shapes
                            ShapeRange pastedShapes2 = app.ActiveWindow.View.Slide.Shapes.Paste();
                            
                            
                            // Try grouping again
                            groupedShape = pastedShapes2.Count > 1 ? 
                                pastedShapes2.Group() : pastedShapes2[1];
                        }
                        catch (Exception innerEx) {
                            MessageBox.Show($"无法组合对象: {innerEx.Message}\n原始错误: {ex.Message}");
                            return;
                        }
                    }

                    // Copy grouped shape
                    groupedShape.Copy();

                    // Delete original shape
                    groupedShape.Delete();

                    // Paste as Enhanced Metafile
                    ShapeRange pastedShapes = app.ActiveWindow.View.Slide.Shapes.PasteSpecial(PpPasteDataType.ppPasteEnhancedMetafile);

                    // // Move to original position
                    // if (pastedShapes != null)
                    // {
                    //     pastedShapes.Left = left;
                    //     pastedShapes.Top = top;
                    //     pastedShapes.Select();
                    // }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"粘贴为增强型图形时出错: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("请先选择要转换的对象。");
            }
        }

        private void imgAutoAlign_rowSpace_TextChanged(object sender, RibbonControlEventArgs e)
        {
            string str1 = imgAutoAlign_rowSpace.Text.Split(new char[] { '≈' })[1];
            if (str1 != null)
            {
                fontSizeEditBox.Text = Regex.Replace(str1, @"[^\d.\d]", "");
            }
            AlignPics();
        }

        private void imgAutoAlign_colNum_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgAutoAlign_colSpace_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgWidthEditBpx_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void imgHeightEditBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            AlignPics();
        }

        private void excludeTextcheckBox_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void excludeTextcheckBox2_Click(object sender, RibbonControlEventArgs e)
        {

        }



        private void donate(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.yuque.com/achuan-2");
        }

        private void developer_website(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.github.com/achuan-2");
        }
    }
}
