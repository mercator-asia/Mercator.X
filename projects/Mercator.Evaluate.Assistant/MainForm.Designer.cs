namespace Mercator.Evaluate.Assistant
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.文件FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OpenShpMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.关闭CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.评定EToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.生成评定表SToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.生成上报分等单元图层BToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.字段FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加土地利用水平评价指标字段UToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.IndexTypeMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.添加产量成本属性AToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助HToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.新建NToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.打开OToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.保存SToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.帮助LToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.OpenShpBGWorker = new System.ComponentModel.BackgroundWorker();
            this.SaveXlsBGWorker = new System.ComponentModel.BackgroundWorker();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.textBox = new System.Windows.Forms.TextBox();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.statusStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件FToolStripMenuItem,
            this.评定EToolStripMenuItem,
            this.字段FToolStripMenuItem,
            this.帮助HToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(818, 25);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "menuStrip1";
            // 
            // 文件FToolStripMenuItem
            // 
            this.文件FToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.OpenShpMenuItem,
            this.toolStripMenuItem1,
            this.关闭CToolStripMenuItem});
            this.文件FToolStripMenuItem.Name = "文件FToolStripMenuItem";
            this.文件FToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.文件FToolStripMenuItem.Text = "文件(&F)";
            // 
            // OpenShpMenuItem
            // 
            this.OpenShpMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("OpenShpMenuItem.Image")));
            this.OpenShpMenuItem.Name = "OpenShpMenuItem";
            this.OpenShpMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.OpenShpMenuItem.Size = new System.Drawing.Size(246, 22);
            this.OpenShpMenuItem.Text = "打开评价单元图层(&O)...";
            this.OpenShpMenuItem.Click += new System.EventHandler(this.OpenShpMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(243, 6);
            // 
            // 关闭CToolStripMenuItem
            // 
            this.关闭CToolStripMenuItem.Name = "关闭CToolStripMenuItem";
            this.关闭CToolStripMenuItem.Size = new System.Drawing.Size(246, 22);
            this.关闭CToolStripMenuItem.Text = "关闭评价单元图层(&C)";
            this.关闭CToolStripMenuItem.Click += new System.EventHandler(this.CloseShpMenuItem_Click);
            // 
            // 评定EToolStripMenuItem
            // 
            this.评定EToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.生成评定表SToolStripMenuItem,
            this.toolStripMenuItem3,
            this.生成上报分等单元图层BToolStripMenuItem});
            this.评定EToolStripMenuItem.Name = "评定EToolStripMenuItem";
            this.评定EToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.评定EToolStripMenuItem.Text = "评定(&E)";
            // 
            // 生成评定表SToolStripMenuItem
            // 
            this.生成评定表SToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("生成评定表SToolStripMenuItem.Image")));
            this.生成评定表SToolStripMenuItem.Name = "生成评定表SToolStripMenuItem";
            this.生成评定表SToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.生成评定表SToolStripMenuItem.Text = "生成评定结果表(&S)";
            this.生成评定表SToolStripMenuItem.Click += new System.EventHandler(this.SaveXlsMenuItem_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(209, 6);
            // 
            // 生成上报分等单元图层BToolStripMenuItem
            // 
            this.生成上报分等单元图层BToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("生成上报分等单元图层BToolStripMenuItem.Image")));
            this.生成上报分等单元图层BToolStripMenuItem.Name = "生成上报分等单元图层BToolStripMenuItem";
            this.生成上报分等单元图层BToolStripMenuItem.Size = new System.Drawing.Size(212, 22);
            this.生成上报分等单元图层BToolStripMenuItem.Text = "生成上报分等单元图层(&B)";
            this.生成上报分等单元图层BToolStripMenuItem.Click += new System.EventHandler(this.SaveLayerAsMenuItem_Click);
            // 
            // 字段FToolStripMenuItem
            // 
            this.字段FToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加土地利用水平评价指标字段UToolStripMenuItem,
            this.toolStripMenuItem2,
            this.IndexTypeMenuItem,
            this.添加产量成本属性AToolStripMenuItem});
            this.字段FToolStripMenuItem.Name = "字段FToolStripMenuItem";
            this.字段FToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.字段FToolStripMenuItem.Text = "设置(&F)";
            // 
            // 添加土地利用水平评价指标字段UToolStripMenuItem
            // 
            this.添加土地利用水平评价指标字段UToolStripMenuItem.Name = "添加土地利用水平评价指标字段UToolStripMenuItem";
            this.添加土地利用水平评价指标字段UToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.添加土地利用水平评价指标字段UToolStripMenuItem.Text = "添加 评定辅助 字段(&U)";
            this.添加土地利用水平评价指标字段UToolStripMenuItem.Click += new System.EventHandler(this.AddCustomFieldsMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(205, 6);
            // 
            // IndexTypeMenuItem
            // 
            this.IndexTypeMenuItem.Checked = true;
            this.IndexTypeMenuItem.CheckOnClick = true;
            this.IndexTypeMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.IndexTypeMenuItem.Name = "IndexTypeMenuItem";
            this.IndexTypeMenuItem.Size = new System.Drawing.Size(208, 22);
            this.IndexTypeMenuItem.Text = "使用气候指数计算水浇地";
            this.IndexTypeMenuItem.CheckedChanged += new System.EventHandler(this.IndexTypeMenuItem_CheckedChanged);
            // 
            // 添加产量成本属性AToolStripMenuItem
            // 
            this.添加产量成本属性AToolStripMenuItem.Name = "添加产量成本属性AToolStripMenuItem";
            this.添加产量成本属性AToolStripMenuItem.Size = new System.Drawing.Size(208, 22);
            this.添加产量成本属性AToolStripMenuItem.Text = "添加 产量/成本 属性(&A)";
            this.添加产量成本属性AToolStripMenuItem.Click += new System.EventHandler(this.SetAttributeMenuItem_Click);
            // 
            // 帮助HToolStripMenuItem
            // 
            this.帮助HToolStripMenuItem.Name = "帮助HToolStripMenuItem";
            this.帮助HToolStripMenuItem.Size = new System.Drawing.Size(61, 21);
            this.帮助HToolStripMenuItem.Text = "帮助(&H)";
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.新建NToolStripButton,
            this.打开OToolStripButton,
            this.保存SToolStripButton,
            this.toolStripSeparator,
            this.帮助LToolStripButton});
            this.toolStrip.Location = new System.Drawing.Point(0, 25);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(818, 25);
            this.toolStrip.TabIndex = 1;
            this.toolStrip.Text = "toolStrip1";
            // 
            // 新建NToolStripButton
            // 
            this.新建NToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.新建NToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("新建NToolStripButton.Image")));
            this.新建NToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.新建NToolStripButton.Name = "新建NToolStripButton";
            this.新建NToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.新建NToolStripButton.Text = "新建(&N)";
            // 
            // 打开OToolStripButton
            // 
            this.打开OToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.打开OToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("打开OToolStripButton.Image")));
            this.打开OToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.打开OToolStripButton.Name = "打开OToolStripButton";
            this.打开OToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.打开OToolStripButton.Text = "打开(&O)";
            this.打开OToolStripButton.Click += new System.EventHandler(this.OpenShpMenuItem_Click);
            // 
            // 保存SToolStripButton
            // 
            this.保存SToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.保存SToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("保存SToolStripButton.Image")));
            this.保存SToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.保存SToolStripButton.Name = "保存SToolStripButton";
            this.保存SToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.保存SToolStripButton.Text = "保存(&S)";
            this.保存SToolStripButton.Click += new System.EventHandler(this.SaveLayerAsMenuItem_Click);
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // 帮助LToolStripButton
            // 
            this.帮助LToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.帮助LToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("帮助LToolStripButton.Image")));
            this.帮助LToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.帮助LToolStripButton.Name = "帮助LToolStripButton";
            this.帮助LToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.帮助LToolStripButton.Text = "帮助(&L)";
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2});
            this.statusStrip.Location = new System.Drawing.Point(0, 428);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(818, 22);
            this.statusStrip.TabIndex = 2;
            this.statusStrip.Text = "statusStrip1";
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.Location = new System.Drawing.Point(0, 0);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowTemplate.Height = 23;
            this.dataGridView.Size = new System.Drawing.Size(818, 289);
            this.dataGridView.TabIndex = 3;
            // 
            // OpenShpBGWorker
            // 
            this.OpenShpBGWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.OpenShpBGWorker_DoWork);
            // 
            // SaveXlsBGWorker
            // 
            this.SaveXlsBGWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.SaveXlsBGWorker_DoWork);
            this.SaveXlsBGWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.SaveXlsBGWorker_RunWorkerCompleted);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 50);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridView);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.textBox);
            this.splitContainer1.Size = new System.Drawing.Size(818, 378);
            this.splitContainer1.SplitterDistance = 289;
            this.splitContainer1.TabIndex = 4;
            // 
            // textBox
            // 
            this.textBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox.Location = new System.Drawing.Point(0, 0);
            this.textBox.Multiline = true;
            this.textBox.Name = "textBox";
            this.textBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox.Size = new System.Drawing.Size(818, 85);
            this.textBox.TabIndex = 0;
            this.textBox.Text = resources.GetString("textBox.Text");
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(28, 17);
            this.toolStripStatusLabel1.Text = "Ver";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(131, 17);
            this.toolStripStatusLabel2.Text = "toolStripStatusLabel2";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(818, 450);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.menuStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "土地整治项目耕地质量等别评定-昆明墨卡托地理信息技术有限公司";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem 文件FToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem OpenShpMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripButton 新建NToolStripButton;
        private System.Windows.Forms.ToolStripButton 打开OToolStripButton;
        private System.Windows.Forms.ToolStripButton 保存SToolStripButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripButton 帮助LToolStripButton;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem 关闭CToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 评定EToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 生成评定表SToolStripMenuItem;
        private System.ComponentModel.BackgroundWorker OpenShpBGWorker;
        private System.ComponentModel.BackgroundWorker SaveXlsBGWorker;
        private System.Windows.Forms.ToolStripMenuItem 字段FToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 添加产量成本属性AToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.ToolStripMenuItem 生成上报分等单元图层BToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助HToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem 添加土地利用水平评价指标字段UToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem IndexTypeMenuItem;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
    }
}

