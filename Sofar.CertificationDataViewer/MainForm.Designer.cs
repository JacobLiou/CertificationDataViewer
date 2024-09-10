namespace Sofar.CertificationDataViewer;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
        DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
        DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
        DataGridViewCellStyle dataGridViewCellStyle4 = new DataGridViewCellStyle();
        DataGridViewCellStyle dataGridViewCellStyle5 = new DataGridViewCellStyle();
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
        textBoxExcelPath = new TextBox();
        buttonImport = new Button();
        buttonSelectFile = new Button();
        btnRead = new Button();
        uiDataGridView1 = new Sunny.UI.UIDataGridView();
        btnExport = new Button();
        uiContextMenuStrip1 = new Sunny.UI.UIContextMenuStrip();
        toolStripMenuItem1 = new ToolStripMenuItem();
        toolStripMenuItem2 = new ToolStripMenuItem();
        toolStripMenuItem3 = new ToolStripMenuItem();
        label1 = new Label();
        uiSplitContainer1 = new Sunny.UI.UISplitContainer();
        ((System.ComponentModel.ISupportInitialize)uiDataGridView1).BeginInit();
        uiContextMenuStrip1.SuspendLayout();
        uiSplitContainer1.BeginInit();
        uiSplitContainer1.Panel1.SuspendLayout();
        uiSplitContainer1.Panel2.SuspendLayout();
        uiSplitContainer1.SuspendLayout();
        SuspendLayout();
        // 
        // textBoxExcelPath
        // 
        textBoxExcelPath.BackColor = Color.Silver;
        textBoxExcelPath.Location = new Point(106, 14);
        textBoxExcelPath.Name = "textBoxExcelPath";
        textBoxExcelPath.Size = new Size(760, 30);
        textBoxExcelPath.TabIndex = 1;
        // 
        // buttonImport
        // 
        buttonImport.BackColor = Color.FromArgb(255, 224, 192);
        buttonImport.Font = new Font("宋体", 12F, FontStyle.Bold, GraphicsUnit.Point);
        buttonImport.ForeColor = Color.Black;
        buttonImport.Location = new Point(1020, 9);
        buttonImport.Name = "buttonImport";
        buttonImport.Size = new Size(131, 42);
        buttonImport.TabIndex = 2;
        buttonImport.Text = "导入数据库";
        buttonImport.UseVisualStyleBackColor = false;
        buttonImport.Click += buttonImport_Click;
        // 
        // buttonSelectFile
        // 
        buttonSelectFile.BackColor = Color.FromArgb(255, 224, 192);
        buttonSelectFile.Font = new Font("宋体", 12F, FontStyle.Bold, GraphicsUnit.Point);
        buttonSelectFile.ForeColor = Color.Black;
        buttonSelectFile.Location = new Point(884, 8);
        buttonSelectFile.Name = "buttonSelectFile";
        buttonSelectFile.Size = new Size(121, 42);
        buttonSelectFile.TabIndex = 4;
        buttonSelectFile.Text = "选择文件";
        buttonSelectFile.UseVisualStyleBackColor = false;
        buttonSelectFile.Click += buttonSelectFile_Click;
        // 
        // btnRead
        // 
        btnRead.BackColor = Color.FromArgb(255, 224, 192);
        btnRead.Font = new Font("宋体", 12F, FontStyle.Bold, GraphicsUnit.Point);
        btnRead.ForeColor = Color.Black;
        btnRead.Location = new Point(1167, 10);
        btnRead.Name = "btnRead";
        btnRead.Size = new Size(125, 42);
        btnRead.TabIndex = 5;
        btnRead.Text = "查询数据库";
        btnRead.UseVisualStyleBackColor = false;
        btnRead.Click += btnRead_Click_1;
        // 
        // uiDataGridView1
        // 
        uiDataGridView1.AllowUserToAddRows = false;
        uiDataGridView1.AllowUserToDeleteRows = false;
        dataGridViewCellStyle1.BackColor = Color.FromArgb(235, 243, 255);
        uiDataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
        uiDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        uiDataGridView1.BackgroundColor = Color.White;
        uiDataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter;
        dataGridViewCellStyle2.BackColor = Color.FromArgb(80, 160, 255);
        dataGridViewCellStyle2.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        dataGridViewCellStyle2.ForeColor = Color.White;
        dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
        dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
        uiDataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
        uiDataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle3.BackColor = SystemColors.Window;
        dataGridViewCellStyle3.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        dataGridViewCellStyle3.ForeColor = Color.FromArgb(48, 48, 48);
        dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
        dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
        dataGridViewCellStyle3.WrapMode = DataGridViewTriState.False;
        uiDataGridView1.DefaultCellStyle = dataGridViewCellStyle3;
        uiDataGridView1.Dock = DockStyle.Fill;
        uiDataGridView1.EnableHeadersVisualStyles = false;
        uiDataGridView1.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        uiDataGridView1.GridColor = Color.FromArgb(80, 160, 255);
        uiDataGridView1.Location = new Point(0, 0);
        uiDataGridView1.Name = "uiDataGridView1";
        dataGridViewCellStyle4.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle4.BackColor = Color.FromArgb(235, 243, 255);
        dataGridViewCellStyle4.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        dataGridViewCellStyle4.ForeColor = Color.FromArgb(48, 48, 48);
        dataGridViewCellStyle4.SelectionBackColor = Color.FromArgb(80, 160, 255);
        dataGridViewCellStyle4.SelectionForeColor = Color.White;
        dataGridViewCellStyle4.WrapMode = DataGridViewTriState.True;
        uiDataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
        uiDataGridView1.RowHeadersWidth = 51;
        dataGridViewCellStyle5.BackColor = Color.White;
        dataGridViewCellStyle5.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        uiDataGridView1.RowsDefaultCellStyle = dataGridViewCellStyle5;
        uiDataGridView1.RowTemplate.Height = 29;
        uiDataGridView1.SelectedIndex = -1;
        uiDataGridView1.Size = new Size(1860, 827);
        uiDataGridView1.StripeOddColor = Color.FromArgb(235, 243, 255);
        uiDataGridView1.TabIndex = 6;
        // 
        // btnExport
        // 
        btnExport.BackColor = Color.FromArgb(128, 128, 255);
        btnExport.Font = new Font("宋体", 12F, FontStyle.Bold, GraphicsUnit.Point);
        btnExport.ForeColor = Color.Black;
        btnExport.Location = new Point(1320, 10);
        btnExport.Name = "btnExport";
        btnExport.Size = new Size(111, 42);
        btnExport.TabIndex = 7;
        btnExport.Text = "导出数据";
        btnExport.UseVisualStyleBackColor = false;
        btnExport.Click += btnExport_Click;
        // 
        // uiContextMenuStrip1
        // 
        uiContextMenuStrip1.BackColor = Color.FromArgb(243, 249, 255);
        uiContextMenuStrip1.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        uiContextMenuStrip1.ImageScalingSize = new Size(20, 20);
        uiContextMenuStrip1.Items.AddRange(new ToolStripItem[] { toolStripMenuItem1, toolStripMenuItem2, toolStripMenuItem3 });
        uiContextMenuStrip1.Name = "uiContextMenuStrip1";
        uiContextMenuStrip1.Size = new Size(159, 76);
        // 
        // toolStripMenuItem1
        // 
        toolStripMenuItem1.Name = "toolStripMenuItem1";
        toolStripMenuItem1.Size = new Size(158, 24);
        toolStripMenuItem1.Text = "新增数据";
        toolStripMenuItem1.Click += toolStripMenuItem1_Click;
        // 
        // toolStripMenuItem2
        // 
        toolStripMenuItem2.Name = "toolStripMenuItem2";
        toolStripMenuItem2.Size = new Size(158, 24);
        toolStripMenuItem2.Text = "删除数据";
        toolStripMenuItem2.Click += toolStripMenuItem2_Click;
        // 
        // toolStripMenuItem3
        // 
        toolStripMenuItem3.Name = "toolStripMenuItem3";
        toolStripMenuItem3.Size = new Size(158, 24);
        toolStripMenuItem3.Text = "修改数据";
        toolStripMenuItem3.Click += toolStripMenuItem3_Click;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.BackColor = Color.White;
        label1.Font = new Font("宋体", 12F, FontStyle.Bold, GraphicsUnit.Point);
        label1.ForeColor = Color.Tomato;
        label1.Location = new Point(7, 21);
        label1.Name = "label1";
        label1.Size = new Size(93, 20);
        label1.TabIndex = 8;
        label1.Text = "文件路径";
        // 
        // uiSplitContainer1
        // 
        uiSplitContainer1.ArrowColor = Color.Yellow;
        uiSplitContainer1.BarColor = Color.Green;
        uiSplitContainer1.Dock = DockStyle.Fill;
        uiSplitContainer1.IsSplitterFixed = true;
        uiSplitContainer1.Location = new Point(2, 36);
        uiSplitContainer1.MinimumSize = new Size(20, 20);
        uiSplitContainer1.Name = "uiSplitContainer1";
        uiSplitContainer1.Orientation = Orientation.Horizontal;
        // 
        // uiSplitContainer1.Panel1
        // 
        uiSplitContainer1.Panel1.Controls.Add(btnExport);
        uiSplitContainer1.Panel1.Controls.Add(textBoxExcelPath);
        uiSplitContainer1.Panel1.Controls.Add(label1);
        uiSplitContainer1.Panel1.Controls.Add(buttonImport);
        uiSplitContainer1.Panel1.Controls.Add(buttonSelectFile);
        uiSplitContainer1.Panel1.Controls.Add(btnRead);
        // 
        // uiSplitContainer1.Panel2
        // 
        uiSplitContainer1.Panel2.Controls.Add(uiDataGridView1);
        uiSplitContainer1.Size = new Size(1860, 920);
        uiSplitContainer1.SplitterDistance = 82;
        uiSplitContainer1.SplitterWidth = 11;
        uiSplitContainer1.TabIndex = 9;
        // 
        // Form1
        // 
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(1864, 958);
        Controls.Add(uiSplitContainer1);
        Icon = (Icon)resources.GetObject("$this.Icon");
        Name = "Form1";
        Padding = new Padding(2, 36, 2, 2);
        RectColor = Color.DarkSlateGray;
        Resizable = true;
        ShowDragStretch = true;
        Text = "产品认证证书信息管理V1.6";
        TitleColor = Color.DarkSlateGray;
        TitleForeColor = Color.Orange;
        ZoomScaleRect = new Rectangle(19, 19, 1304, 695);
        Load += Form1_Load;
        ((System.ComponentModel.ISupportInitialize)uiDataGridView1).EndInit();
        uiContextMenuStrip1.ResumeLayout(false);
        uiSplitContainer1.Panel1.ResumeLayout(false);
        uiSplitContainer1.Panel1.PerformLayout();
        uiSplitContainer1.Panel2.ResumeLayout(false);
        uiSplitContainer1.EndInit();
        uiSplitContainer1.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion
    private TextBox textBoxExcelPath;
    private Button buttonImport;
    private Button buttonSelectFile;
    private Button btnRead;
    private Sunny.UI.UIDataGridView uiDataGridView1;
    private Button btnExport;
    private Sunny.UI.UIContextMenuStrip uiContextMenuStrip1;
    private ToolStripMenuItem toolStripMenuItem1;
    private ToolStripMenuItem toolStripMenuItem2;
    private ToolStripMenuItem toolStripMenuItem3;
    private Label label1;
    private Sunny.UI.UISplitContainer uiSplitContainer1;
}
