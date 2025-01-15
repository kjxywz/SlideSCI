namespace Achuan的PPT插件
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.图片处理 = this.Factory.CreateRibbonGroup();
            this.AddTitleButton = this.Factory.CreateRibbonButton();
            this.fontNameEditBox = this.Factory.CreateRibbonEditBox();
            this.fontSizeEditBox = this.Factory.CreateRibbonEditBox();
            this.distanceFromBottomEditBox = this.Factory.CreateRibbonEditBox();
            this.titleTextEditBox = this.Factory.CreateRibbonEditBox();
            this.autoGroupCheckBox = this.Factory.CreateRibbonCheckBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.copyPosition = this.Factory.CreateRibbonButton();
            this.pastePosition = this.Factory.CreateRibbonButton();
            this.复制图片格式 = this.Factory.CreateRibbonGroup();
            this.copyImgWidth = this.Factory.CreateRibbonButton();
            this.pasteImgWidth = this.Factory.CreateRibbonButton();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.copyImgHeight = this.Factory.CreateRibbonButton();
            this.pasteImgHeight = this.Factory.CreateRibbonButton();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.copyCrop = this.Factory.CreateRibbonButton();
            this.pasteCrop = this.Factory.CreateRibbonButton();
            this.图片自动对齐 = this.Factory.CreateRibbonGroup();
            this.imgAutoAlign = this.Factory.CreateRibbonButton();
            this.imgAutoAlign_colNum = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_colSpace = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_rowSpace = this.Factory.CreateRibbonEditBox();
            this.imgWidthEditBpx = this.Factory.CreateRibbonEditBox();
            this.imgHeightEditBox = this.Factory.CreateRibbonEditBox();
            this.codeGroup = this.Factory.CreateRibbonGroup();
            this.insertCodeBlockButton = this.Factory.CreateRibbonButton();
            this.toggleBackgroundButton = this.Factory.CreateRibbonToggleButton();
            this.insertEquationButton = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.图片处理.SuspendLayout();
            this.group1.SuspendLayout();
            this.复制图片格式.SuspendLayout();
            this.图片自动对齐.SuspendLayout();
            this.codeGroup.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.图片处理);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.复制图片格式);
            this.tab1.Groups.Add(this.图片自动对齐);
            this.tab1.Groups.Add(this.codeGroup);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "Achuan的插件";
            this.tab1.Name = "tab1";
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
            this.distanceFromBottomEditBox.Text = "5";
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
            this.group1.Items.Add(this.copyPosition);
            this.group1.Items.Add(this.pastePosition);
            this.group1.Label = "复制粘贴位置";
            this.group1.Name = "group1";
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
            // 复制图片格式
            // 
            this.复制图片格式.Items.Add(this.copyImgWidth);
            this.复制图片格式.Items.Add(this.pasteImgWidth);
            this.复制图片格式.Items.Add(this.label1);
            this.复制图片格式.Items.Add(this.copyImgHeight);
            this.复制图片格式.Items.Add(this.pasteImgHeight);
            this.复制图片格式.Items.Add(this.label2);
            this.复制图片格式.Items.Add(this.copyCrop);
            this.复制图片格式.Items.Add(this.pasteCrop);
            this.复制图片格式.Label = "复制图片格式";
            this.复制图片格式.Name = "复制图片格式";
            // 
            // copyImgWidth
            // 
            this.copyImgWidth.Image = ((System.Drawing.Image)(resources.GetObject("copyImgWidth.Image")));
            this.copyImgWidth.Label = "复制图片宽度";
            this.copyImgWidth.Name = "copyImgWidth";
            this.copyImgWidth.ShowImage = true;
            this.copyImgWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyImgWidth_Click);
            // 
            // pasteImgWidth
            // 
            this.pasteImgWidth.Image = ((System.Drawing.Image)(resources.GetObject("pasteImgWidth.Image")));
            this.pasteImgWidth.Label = "粘贴图片宽度";
            this.pasteImgWidth.Name = "pasteImgWidth";
            this.pasteImgWidth.ShowImage = true;
            this.pasteImgWidth.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteImgWidth_Click);
            // 
            // label1
            // 
            this.label1.Label = " ";
            this.label1.Name = "label1";
            // 
            // copyImgHeight
            // 
            this.copyImgHeight.Image = ((System.Drawing.Image)(resources.GetObject("copyImgHeight.Image")));
            this.copyImgHeight.Label = "复制图片高度";
            this.copyImgHeight.Name = "copyImgHeight";
            this.copyImgHeight.ShowImage = true;
            this.copyImgHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyImgHeight_Click);
            // 
            // pasteImgHeight
            // 
            this.pasteImgHeight.Image = ((System.Drawing.Image)(resources.GetObject("pasteImgHeight.Image")));
            this.pasteImgHeight.Label = "粘贴图片高度";
            this.pasteImgHeight.Name = "pasteImgHeight";
            this.pasteImgHeight.ShowImage = true;
            this.pasteImgHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteImgHeight_Click);
            // 
            // label2
            // 
            this.label2.Label = " ";
            this.label2.Name = "label2";
            // 
            // copyCrop
            // 
            this.copyCrop.Label = "复制裁剪";
            this.copyCrop.Name = "copyCrop";
            this.copyCrop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyCrop_Click);
            // 
            // pasteCrop
            // 
            this.pasteCrop.Label = "粘贴裁剪";
            this.pasteCrop.Name = "pasteCrop";
            this.pasteCrop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteCrop_Click);
            // 
            // 图片自动对齐
            // 
            this.图片自动对齐.Items.Add(this.imgAutoAlign);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_colNum);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_colSpace);
            this.图片自动对齐.Items.Add(this.imgAutoAlign_rowSpace);
            this.图片自动对齐.Items.Add(this.imgWidthEditBpx);
            this.图片自动对齐.Items.Add(this.imgHeightEditBox);
            this.图片自动对齐.Label = "图片自动对齐";
            this.图片自动对齐.Name = "图片自动对齐";
            // 
            // imgAutoAlign
            // 
            this.imgAutoAlign.Image = ((System.Drawing.Image)(resources.GetObject("imgAutoAlign.Image")));
            this.imgAutoAlign.Label = "图片自动对齐";
            this.imgAutoAlign.Name = "imgAutoAlign";
            this.imgAutoAlign.ShowImage = true;
            this.imgAutoAlign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.imgAutoAlign_Click);
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
            this.imgWidthEditBpx.Text = null;
            // 
            // imgHeightEditBox
            // 
            this.imgHeightEditBox.Label = "图片高度";
            this.imgHeightEditBox.Name = "imgHeightEditBox";
            this.imgHeightEditBox.Text = null;
            // 
            // codeGroup
            // 
            this.codeGroup.Items.Add(this.insertCodeBlockButton);
            this.codeGroup.Items.Add(this.toggleBackgroundButton);
            this.codeGroup.Items.Add(this.insertEquationButton);
            this.codeGroup.Items.Add(this.button2);
            this.codeGroup.Label = "Markdown";
            this.codeGroup.Name = "codeGroup";
            // 
            // insertCodeBlockButton
            // 
            this.insertCodeBlockButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.insertCodeBlockButton.Image = ((System.Drawing.Image)(resources.GetObject("insertCodeBlockButton.Image")));
            this.insertCodeBlockButton.Label = "插入代码块";
            this.insertCodeBlockButton.Name = "insertCodeBlockButton";
            this.insertCodeBlockButton.ShowImage = true;
            this.insertCodeBlockButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertCodeBlockButton_Click);
            // 
            // toggleBackgroundButton
            // 
            this.toggleBackgroundButton.Checked = true;
            this.toggleBackgroundButton.Label = "黑色背景色";
            this.toggleBackgroundButton.Name = "toggleBackgroundButton";
            this.toggleBackgroundButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleBackgroundButton_Click);
            // 
            // insertEquationButton
            // 
            this.insertEquationButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.insertEquationButton.Image = ((System.Drawing.Image)(resources.GetObject("insertEquationButton.Image")));
            this.insertEquationButton.Label = "插入latex公式";
            this.insertEquationButton.Name = "insertEquationButton";
            this.insertEquationButton.ShowImage = true;
            this.insertEquationButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertEquationButton_Click);
            // 
            // button2
            // 
            this.button2.Label = "插入Markdown";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.insertMarkdown_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.button4);
            this.group2.Label = "关于";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "开发者";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Label = "Github";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.openGithub_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.图片处理.ResumeLayout(false);
            this.图片处理.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.复制图片格式.ResumeLayout(false);
            this.复制图片格式.PerformLayout();
            this.图片自动对齐.ResumeLayout(false);
            this.图片自动对齐.PerformLayout();
            this.codeGroup.ResumeLayout(false);
            this.codeGroup.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTitleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyPosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pastePosition;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton imgAutoAlign;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_colNum;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_colSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgAutoAlign_rowSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片自动对齐;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgWidthEditBpx;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox imgHeightEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyImgWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteImgWidth;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyImgHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteImgHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 复制图片格式;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox fontSizeEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox distanceFromBottomEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox autoGroupCheckBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox fontNameEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox titleTextEditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertCodeBlockButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup codeGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleBackgroundButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton insertEquationButton; // Add this line
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyCrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteCrop;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
