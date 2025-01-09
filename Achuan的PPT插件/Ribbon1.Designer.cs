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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.图片自动对齐 = this.Factory.CreateRibbonGroup();
            this.imgAutoAlign_colNum = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_colSpace = this.Factory.CreateRibbonEditBox();
            this.imgAutoAlign_rowSpace = this.Factory.CreateRibbonEditBox();
            this.imgWidthEditBpx = this.Factory.CreateRibbonEditBox();
            this.imgHeightEditBox = this.Factory.CreateRibbonEditBox();
            this.AddTitleButton = this.Factory.CreateRibbonButton();
            this.copyPosition = this.Factory.CreateRibbonButton();
            this.pastePosition = this.Factory.CreateRibbonButton();
            this.copyImgWidthHeight = this.Factory.CreateRibbonButton();
            this.pasteImgWidthHeight = this.Factory.CreateRibbonButton();
            this.imgAutoAlign = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.图片处理.SuspendLayout();
            this.group1.SuspendLayout();
            this.图片自动对齐.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.图片处理);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.图片自动对齐);
            this.tab1.Label = "Achuan的插件";
            this.tab1.Name = "tab1";
            // 
            // 图片处理
            // 
            this.图片处理.Items.Add(this.AddTitleButton);
            this.图片处理.Name = "图片处理";
            // 
            // group1
            // 
            this.group1.Items.Add(this.copyPosition);
            this.group1.Items.Add(this.pastePosition);
            this.group1.Items.Add(this.copyImgWidthHeight);
            this.group1.Items.Add(this.pasteImgWidthHeight);
            this.group1.Label = "样式统一";
            this.group1.Name = "group1";
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
            // copyImgWidthHeight
            // 
            this.copyImgWidthHeight.Image = ((System.Drawing.Image)(resources.GetObject("copyImgWidthHeight.Image")));
            this.copyImgWidthHeight.Label = "复制图片宽高";
            this.copyImgWidthHeight.Name = "copyImgWidthHeight";
            this.copyImgWidthHeight.ShowImage = true;
            this.copyImgWidthHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.copyImgWidthHeight_Click);
            // 
            // pasteImgWidthHeight
            // 
            this.pasteImgWidthHeight.Image = ((System.Drawing.Image)(resources.GetObject("pasteImgWidthHeight.Image")));
            this.pasteImgWidthHeight.Label = "粘贴图片宽高";
            this.pasteImgWidthHeight.Name = "pasteImgWidthHeight";
            this.pasteImgWidthHeight.ShowImage = true;
            this.pasteImgWidthHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.pasteImgWidthHeight_Click);
            // 
            // imgAutoAlign
            // 
            this.imgAutoAlign.Image = ((System.Drawing.Image)(resources.GetObject("imgAutoAlign.Image")));
            this.imgAutoAlign.Label = "图片自动对齐";
            this.imgAutoAlign.Name = "imgAutoAlign";
            this.imgAutoAlign.ShowImage = true;
            this.imgAutoAlign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.imgAutoAlign_Click);
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
            this.图片自动对齐.ResumeLayout(false);
            this.图片自动对齐.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup 图片处理;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTitleButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton copyImgWidthHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton pasteImgWidthHeight;
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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
