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
        uiDataGridView1 = new Sunny.UI.UIDataGridView();
        ((System.ComponentModel.ISupportInitialize)uiDataGridView1).BeginInit();
        SuspendLayout();
        // 
        // uiDataGridView1
        // 
        dataGridViewCellStyle1.BackColor = Color.FromArgb(235, 243, 255);
        uiDataGridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
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
        uiDataGridView1.EnableHeadersVisualStyles = false;
        uiDataGridView1.Font = new Font("宋体", 12F, FontStyle.Regular, GraphicsUnit.Point);
        uiDataGridView1.GridColor = Color.FromArgb(80, 160, 255);
        uiDataGridView1.Location = new Point(131, 182);
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
        uiDataGridView1.Size = new Size(965, 352);
        uiDataGridView1.StripeOddColor = Color.FromArgb(235, 243, 255);
        uiDataGridView1.TabIndex = 0;
        // 
        // Form1
        // 
        AutoScaleMode = AutoScaleMode.None;
        ClientSize = new Size(1304, 695);
        Controls.Add(uiDataGridView1);
        Name = "Form1";
        Text = "认证数据查看器";
        ZoomScaleRect = new Rectangle(19, 19, 1304, 695);
        ((System.ComponentModel.ISupportInitialize)uiDataGridView1).EndInit();
        ResumeLayout(false);
    }

    #endregion

    private Sunny.UI.UIDataGridView uiDataGridView1;
}
