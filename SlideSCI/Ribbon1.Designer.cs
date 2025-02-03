namespace SlideSCI
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.图片自动对齐 = this.Factory.CreateRibbonGroup();
            this.imgAutoAlign = this.Factory.CreateRibbonButton();
            this.imgAutoAlignSortTypeDropDown = this.Factory.CreateRibbonDropDown();
            this.imgAutoAlignAlignTypeDropDown = this.Factory.CreateRibbonDropDown();
            this.excludeTextcheckBox = this.Factory.CreateRibbonCheckBox();
            this.imgAutoAlign_colNum = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_colSpace = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_rowSpace = this.Factory.CreateRibbonEditBox();
            this.imgWidthEditBpx = this.Factory.CreateRibbonEditBox();
            this.imgHeightEditBox = this.Factory.CreateRibbonEditBox();
            this.图片处理 = this.Factory.CreateRibbonGroup();
            this.AddTitleButton = this.Factory.CreateRibbonButton();
            this.fontNameEditBox = this.Factory.CreateRibbonEditBox();
            this.fontSizeEditBox = this.Factory.CreateRibbonEditBox();
            this.distanceFromBottomEditBox = this.Factory.CreateRibbonEditBox();
            this.titleTextEditBox = this.Factory.CreateRibbonEditBox();
            this.autoGroupCheckBox = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.addLabelsButton = this.Factory.CreateRibbonButton();
            this.labelFontSizeEditBox = this.Factory.CreateRibbonEditBox();
            this.labelFontNameEditBox = this.Factory.CreateRibbonEditBox();
            this.labelTemplateComboBox = this.Factory.CreateRibbonComboBox();
            this.labelOffsetYEditBox = this.Factory.CreateRibbonEditBox();
            this.labelOffsetXEditBox = this.Factory.CreateRibbonEditBox();
            this.labelBoldcheckBox = this.Factory.CreateRibbonCheckBox();
            this.复制图片格式 = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.copyPosition = this.Factory.CreateRibbonButton();
            this.pastePosition = this.Factory.CreateRibbonButton();
            this.copyImgWidth = this.Factory.CreateRibbonButton();
            this.pasteImgWidth = this.Factory.CreateRibbonButton();
            this.copyImgHeight = this.Factory.CreateRibbonButton();
            this.pasteImgHeight = this.Factory.CreateRibbonButton();
            this.copyCrop = this.Factory.CreateRibbonButton();
            this.pasteCrop = this.Factory.CreateRibbonButton();
            this.label3 = this.Factory.CreateRibbonLabel();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.codeGroup = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.insertEquationButton = this.Factory.CreateRibbonButton();
            this.insertCodeBlockButton = this.Factory.CreateRibbonButton();
            this.toggleBackgroundCheckBox = this.Factory.CreateRibbonCheckBox();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.图片自动对齐.SuspendLayout();
            this.图片处理.SuspendLayout();
            this.group1.SuspendLayout();
            this.复制图片格式.SuspendLayout();
            this.codeGroup.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.图片自动对齐);
            this.tab2.Groups.Add(this.图片处理);
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.复制图片格式);
            this.tab2.Groups.Add(this.codeGroup);
            this.tab2.Groups.Add(this.group3);
            this.tab2.Groups.Add(this.group2);
            this.tab2.Label = "SlideSCI";
            this.tab2.Name = "tab2";
            // 
            // 图片自动对齐
            // 
            this.图片自动对齐.Items.Add(this.imgAutoAlign);
            this.图片自动对齐.Items.Add(this.imgAutoAlignSortTypeDropDown);
            this.图片自动对齐.Items.Add(this.imgAutoAlignAlignTypeDropDown);
            this.图片自动对齐.Items.Add(this.excludeTextcheckBox);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_colNum);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_colSpace);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_rowSpace);
            this.图片自动对齐.Items.Add(this.imgWidthEditBpx);
            this.图片自动对齐.Items.Add(this.imgHeightEditBox);
            this.图片自动对齐.Label = "图片自动排列";
            this.图片自动对齐.Name = "图片自动对齐";
            // 
            // imgAutoAlign
            // 
            this.imgAutoAlign.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.imgAutoAlign.Image = ((System.Drawing.Image)(resources.GetObject("imgAutoAlign.Image")));
            this.imgAutoAlign.Label = "图片自动排列";
            this.imgAutoAlign.Name = "imgAutoAlign";
            this.imgAutoAlign.ShowImage = true;
            this.imgAutoAlign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.imgAutoAlign_Click);
            // 
            // imgAutoAlignSortTypeDropDown
            // 
            ribbonDropDownItemImpl1.Label = "根据位置排序";
            ribbonDropDownItemImpl1.ScreenTip = "用户可以先大概行列排列好，程序自动获取正确顺序";
            ribbonDropDownItemImpl2.Label = "根据多选顺序排序";
            ribbonDropDownItemImpl2.ScreenTip = "根据用户多选的顺序来获取图片排列的顺序";
            this.imgAutoAlignSortTypeDropDown.Items.Add(ribbonDropDownItemImpl1);
            this.imgAutoAlignSortTypeDropDown.Items.Add(ribbonDropDownItemImpl2);
            this.imgAutoAlignSortTypeDropDown.Label = "排序方式";
            this.imgAutoAlignSortTypeDropDown.Name = "imgAutoAlignSortTypeDropDown";
            // 
            // imgAutoAlignAlignTypeDropDown
            // 
            ribbonDropDownItemImpl3.Label = "列最大宽度占位排列";
            ribbonDropDownItemImpl3.ScreenTip = "按每列的最大宽度来占位排列，以保持表格布局";
            ribbonDropDownItemImpl4.Label = "统一高度排列";
            ribbonDropDownItemImpl4.ScreenTip = "默认会统一图片的高度整齐紧凑排列在一起";
            ribbonDropDownItemImpl5.Label = "统一宽度瀑布流";
            ribbonDropDownItemImpl5.ScreenTip = "默认图片统一宽度紧凑排列";
            this.imgAutoAlignAlignTypeDropDown.Items.Add(ribbonDropDownItemImpl3);
            this.imgAutoAlignAlignTypeDropDown.Items.Add(ribbonDropDownItemImpl4);
            this.imgAutoAlignAlignTypeDropDown.Items.Add(ribbonDropDownItemImpl5);
            this.imgAutoAlignAlignTypeDropDown.Label = "排列方式";
            this.imgAutoAlignAlignTypeDropDown.Name = "imgAutoAlignAlignTypeDropDown";
            // 
            // excludeTextcheckBox
            // 
            this.excludeTextcheckBox.Checked = true;
            this.excludeTextcheckBox.Label = "排除文本框和形状";
            this.excludeTextcheckBox.Name = "excludeTextcheckBox";
            // 
            // imgAutoAlign_colNum
            // 
            this.imgAutoAlign_colNum.Label = "列数量";
            this.imgAutoAlign_colNum.Name = "imgAutoAlign_colNum";
            this.imgAutoAlign_colNum.Text = "5";
            // 
            // imgAutoAlign_colSpace
            // 
            this.imgAutoAlign_colSpace.Label = "列间距";
            this.imgAutoAlign_colSpace.Name = "imgAutoAlign_colSpace";
            this.imgAutoAlign_colSpace.Text = "5";
            // 
            // imgAutoAlign_rowSpace
            // 
            this.imgAutoAlign_rowSpace.Label = "行间距";
            this.imgAutoAlign_rowSpace.Name = "imgAutoAlign_rowSpace";
            this.imgAutoAlign_rowSpace.Text = null;
            // 
            // imgWidthEditBpx
            // 
            this.imgWidthEditBpx.Label = "图片宽度";
            this.imgWidthEditBpx.Name = "imgWidthEditBpx";
            this.imgWidthEditBpx.ScreenTip = "统一设置图片宽度";
            this.imgWidthEditBpx.Text = null;
            // 
            // imgHeightEditBox
            // 
            this.imgHeightEditBox.Label = "图片高度";
            this.imgHeightEditBox.Name = "imgHeightEditBox";
            this.imgHeightEditBox.ScreenTip = "统一设置图片高度";
            this.imgHeightEditBox.Text = null;
            // 
            // 图片处理
            // 
            this.图片处理.Items.Add(this.AddTitleButton);
            this.图片处理.Items.Add(this.fontNameEditBox);
            this.图片处理.Items.Add(this.fontSizeEditBox);
            this.图片处理.Items.Add(this.distanceFromBottomEditBox);
            this.图片处理.Items.Add(this.titleTextEditBox);
            this.图片处理.Items.Add(this.autoGroupCheckBox);
            this.图片处理.Label = "添加图片标题";
            this.图片处理.Name = "图片处理";
            // 
            // AddTitleButton
            // 
            this.AddTitleButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddTitleButton.Description = "添加图片标题";
            this.AddTitleButton.Image = ((System.Drawing.Image)(resources.GetObject("AddTitleButton.Image")));
            this.AddTitleButton.Label = "添加图片标题";
            this.AddTitleButton.Name = "AddTitleButton";
            this.AddTitleButton.ShowImage = true;
            this.AddTitleButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddTitleToImage);
            // 
            // fontNameEditBox
            // 
            this.fontNameEditBox.Label = "字体名称";
            this.fontNameEditBox.Name = "fontNameEditBox";
            this.fontNameEditBox.Text = "微软雅黑";
            // 
            // fontSizeEditBox
            // 
            this.fontSizeEditBox.Label = "字体大小";
            this.fontSizeEditBox.Name = "fontSizeEditBox";
            this.fontSizeEditBox.Text = "14";
            // 
            // distanceFromBottomEditBox
            // 
            this.distanceFromBottomEditBox.Label = "距离图片下边距离";
            this.distanceFromBottomEditBox.Name = "distanceFromBottomEditBox";
            this.distanceFromBottomEditBox.Text = "0";
            // 
            // titleTextEditBox
            // 
            this.titleTextEditBox.Label = "标题文本";
            this.titleTextEditBox.Name = "titleTextEditBox";
            this.titleTextEditBox.Text = "图片标题";
            // 
            // autoGroupCheckBox
            // 
            this.autoGroupCheckBox.Label = "自动编组";
            this.autoGroupCheckBox.Name = "autoGroupCheckBox";
            // 
            // group1
            // 
            this.group1.Items.Add(this.addLabelsButton);
            this.group1.Items.Add(this.labelFontSizeEditBox);
            this.group1.Items.Add(this.labelFontNameEditBox);
            this.group1.Items.Add(this.labelTemplateComboBox);
            this.group1.Items.Add(this.labelOffsetYEditBox);
            this.group1.Items.Add(this.labelOffsetXEditBox);
            this.group1.Items.Add(this.labelBoldcheckBox);
            this.group1.Label = "添加图片标签";
            this.group1.Name = "group1";
            // 
            // addLabelsButton
            // 
            this.addLabelsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.addLabelsButton.Image = ((System.Drawing.Image)(resources.GetObject("addLabelsButton.Image")));
            this.addLabelsButton.Label = "添加图片标签";
            this.addLabelsButton.Name = "addLabelsButton";
            this.addLabelsButton.ShowImage = true;
            this.addLabelsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addLabelsButton_Click);
            // 
            // labelFontSizeEditBox
            // 
            this.labelFontSizeEditBox.Label = "标签字号";
            this.labelFontSizeEditBox.Name = "labelFontSizeEditBox";
            this.labelFontSizeEditBox.Text = "12";
            // 
            // labelFontNameEditBox
            // 
            this.labelFontNameEditBox.Label = "标签字体";
            this.labelFontNameEditBox.Name = "labelFontNameEditBox";
            this.labelFontNameEditBox.Text = "Arial";
            // 
            // labelTemplateComboBox
            // 
            ribbonDropDownItemImpl6.Label = "A";
            ribbonDropDownItemImpl7.Label = "a";
            ribbonDropDownItemImpl8.Label = "A)";
            ribbonDropDownItemImpl9.Label = "a)";
            ribbonDropDownItemImpl10.Label = "1";
            ribbonDropDownItemImpl11.Label = "1)";
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl6);
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl7);
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl8);
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl9);
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl10);
            this.labelTemplateComboBox.Items.Add(ribbonDropDownItemImpl11);
            this.labelTemplateComboBox.Label = "标签模板";
            this.labelTemplateComboBox.Name = "labelTemplateComboBox";
            this.labelTemplateComboBox.Text = "A";
            // 
            // labelOffsetYEditBox
            // 
            this.labelOffsetYEditBox.Label = "Y Offset";
            this.labelOffsetYEditBox.Name = "labelOffsetYEditBox";
            this.labelOffsetYEditBox.Text = "-7";
            // 
            // labelOffsetXEditBox
            // 
            this.labelOffsetXEditBox.Label = "X Offset";
            this.labelOffsetXEditBox.Name = "labelOffsetXEditBox";
            this.labelOffsetXEditBox.Text = "-20";
            // 
            // labelBoldcheckBox
            // 
            this.labelBoldcheckBox.Checked = true;
            this.labelBoldcheckBox.Label = "自动加粗";
            this.labelBoldcheckBox.Name = "labelBoldcheckBox";
            // 
            // 复制图片格式
            // 
            this.复制图片格式.Items.Add(this.menu1);
            this.复制图片格式.Items.Add(this.label3);
            this.复制图片格式.Items.Add(this.label1);
            this.复制图片格式.Items.Add(this.label2);
            this.复制图片格式.Label = "复制格式";
            this.复制图片格式.Name = "复制图片格式";
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Image = ((System.Drawing.Image)(resources.GetObject("menu1.Image")));
            this.menu1.Items.Add(this.button6);
            this.menu1.Items.Add(this.button7);
            this.menu1.Items.Add(this.copyPosition);
            this.menu1.Items.Add(this.pastePosition);
            this.menu1.Items.Add(this.copyImgWidth);
            this.menu1.Items.Add(this.pasteImgWidth);
            this.menu1.Items.Add(this.copyImgHeight);
            this.menu1.Items.Add(this.pasteImgHeight);
            this.menu1.Items.Add(this.copyCrop);
            this.menu1.Items.Add(this.pasteCrop);
            this.menu1.Label = "复制粘贴格式";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "复制格式";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyStyle_Click);
            // 
            // button7
            // 
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "粘贴格式";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteStyle_Click);
            // 
            // copyPosition
            // 
            this.copyPosition.Image = ((System.Drawing.Image)(resources.GetObject("copyPosition.Image")));
            this.copyPosition.Label = "复制位置";
            this.copyPosition.Name = "copyPosition";
            this.copyPosition.ShowImage = true;
            this.copyPosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyPosition_Click);
            // 
            // pastePosition
            // 
            this.pastePosition.Image = ((System.Drawing.Image)(resources.GetObject("pastePosition.Image")));
            this.pastePosition.Label = "粘贴位置";
            this.pastePosition.Name = "pastePosition";
            this.pastePosition.ShowImage = true;
            this.pastePosition.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pastePosition_Click);
            // 
            // copyImgWidth
            // 
            this.copyImgWidth.Image = ((System.Drawing.Image)(resources.GetObject("copyImgWidth.Image")));
            this.copyImgWidth.Label = "复制宽度";
            this.copyImgWidth.Name = "copyImgWidth";
            this.copyImgWidth.ShowImage = true;
            this.copyImgWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyImgWidth_Click);
            // 
            // pasteImgWidth
            // 
            this.pasteImgWidth.Image = ((System.Drawing.Image)(resources.GetObject("pasteImgWidth.Image")));
            this.pasteImgWidth.Label = "粘贴宽度";
            this.pasteImgWidth.Name = "pasteImgWidth";
            this.pasteImgWidth.ShowImage = true;
            this.pasteImgWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteImgWidth_Click);
            // 
            // copyImgHeight
            // 
            this.copyImgHeight.Image = ((System.Drawing.Image)(resources.GetObject("copyImgHeight.Image")));
            this.copyImgHeight.Label = "复制高度";
            this.copyImgHeight.Name = "copyImgHeight";
            this.copyImgHeight.ShowImage = true;
            this.copyImgHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyImgHeight_Click);
            // 
            // pasteImgHeight
            // 
            this.pasteImgHeight.Image = ((System.Drawing.Image)(resources.GetObject("pasteImgHeight.Image")));
            this.pasteImgHeight.Label = "粘贴高度";
            this.pasteImgHeight.Name = "pasteImgHeight";
            this.pasteImgHeight.ShowImage = true;
            this.pasteImgHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteImgHeight_Click);
            // 
            // copyCrop
            // 
            this.copyCrop.Image = ((System.Drawing.Image)(resources.GetObject("copyCrop.Image")));
            this.copyCrop.Label = "复制图片裁剪";
            this.copyCrop.Name = "copyCrop";
            this.copyCrop.ShowImage = true;
            this.copyCrop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyCrop_Click);
            // 
            // pasteCrop
            // 
            this.pasteCrop.Image = ((System.Drawing.Image)(resources.GetObject("pasteCrop.Image")));
            this.pasteCrop.Label = "粘贴图片裁剪";
            this.pasteCrop.Name = "pasteCrop";
            this.pasteCrop.ShowImage = true;
            this.pasteCrop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteCrop_Click);
            // 
            // label3
            // 
            this.label3.Label = " ";
            this.label3.Name = "label3";
            // 
            // label1
            // 
            this.label1.Label = " ";
            this.label1.Name = "label1";
            // 
            // label2
            // 
            this.label2.Label = " ";
            this.label2.Name = "label2";
            // 
            // codeGroup
            // 
            this.codeGroup.Items.Add(this.button2);
            this.codeGroup.Items.Add(this.insertEquationButton);
            this.codeGroup.Items.Add(this.insertCodeBlockButton);
            this.codeGroup.Items.Add(this.toggleBackgroundCheckBox);
            this.codeGroup.Label = "Markdown";
            this.codeGroup.Name = "codeGroup";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "插入Markdown";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertMarkdown_Click);
            // 
            // insertEquationButton
            // 
            this.insertEquationButton.Image = ((System.Drawing.Image)(resources.GetObject("insertEquationButton.Image")));
            this.insertEquationButton.Label = "插入latex公式";
            this.insertEquationButton.Name = "insertEquationButton";
            this.insertEquationButton.ShowImage = true;
            this.insertEquationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertEquationButton_Click);
            // 
            // insertCodeBlockButton
            // 
            this.insertCodeBlockButton.Image = ((System.Drawing.Image)(resources.GetObject("insertCodeBlockButton.Image")));
            this.insertCodeBlockButton.Label = "插入代码块";
            this.insertCodeBlockButton.Name = "insertCodeBlockButton";
            this.insertCodeBlockButton.ShowImage = true;
            this.insertCodeBlockButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertCodeBlockButton_Click);
            // 
            // toggleBackgroundCheckBox
            // 
            this.toggleBackgroundCheckBox.Checked = true;
            this.toggleBackgroundCheckBox.Label = "代码黑色背景";
            this.toggleBackgroundCheckBox.Name = "toggleBackgroundCheckBox";
            // 
            // group2
            // 
            this.group2.Items.Add(this.splitButton1);
            this.group2.Label = "关于";
            this.group2.Name = "group2";
            // 
            // button4
            // 
            this.button4.Label = "Github";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openGithub_Click);
            // 
            // button3
            // 
            this.button3.Label = "使用介绍";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openDoc_Click);
            // 
            // button5
            // 
            this.button5.Label = "当前版本";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.current_Version);
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.button4);
            this.splitButton1.Items.Add(this.button3);
            this.splitButton1.Items.Add(this.button5);
            this.splitButton1.Label = "开发者";
            this.splitButton1.Name = "splitButton1";
            this.splitButton1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.aboutDeveloper_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button1);
            this.group3.Label = "其他";
            this.group3.Name = "group3";
            // 
            // button1
            // 
            this.button1.Label = "图片和文字同步缩放";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "组合后，调整组合形状，文字也会同步缩放";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pastePictureAndText);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.图片自动对齐.ResumeLayout(false);
            this.图片自动对齐.PerformLayout();
            this.图片处理.ResumeLayout(false);
            this.图片处理.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.复制图片格式.ResumeLayout(false);
            this.复制图片格式.PerformLayout();
            this.codeGroup.ResumeLayout(false);
            this.codeGroup.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片自动对齐;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton imgAutoAlign;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_colNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_colSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_rowSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgWidthEditBpx;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgHeightEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTitleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox fontNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox fontSizeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox distanceFromBottomEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox titleTextEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoGroupCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addLabelsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox labelOffsetXEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox labelOffsetYEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox labelTemplateComboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox labelFontNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox labelFontSizeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 复制图片格式;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteImgWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pastePosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyImgWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyImgHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteImgHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyCrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteCrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup codeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertEquationButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertCodeBlockButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox toggleBackgroundCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown imgAutoAlignSortTypeDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown imgAutoAlignAlignTypeDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox excludeTextcheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox labelBoldcheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
