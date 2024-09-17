
namespace Save2MdbTool
{
    partial class FormSave2MdbTool
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.textBoxCmd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.splitContainer8 = new System.Windows.Forms.SplitContainer();
            this.cbx_Model = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonSearch = new System.Windows.Forms.Button();
            this.textBox_Code = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.dataGridViewDynamic1 = new MesDatasCore.DataGridViewDynamic();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer8)).BeginInit();
            this.splitContainer8.Panel1.SuspendLayout();
            this.splitContainer8.Panel2.SuspendLayout();
            this.splitContainer8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDynamic1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "本地数据库路径";
            this.label1.Click += new System.EventHandler(this.Label1_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(12, 48);
            this.textBoxPath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(860, 23);
            this.textBoxPath.TabIndex = 1;
            this.textBoxPath.Text = "D:\\\\0000000\\\\2021年11月生产数据.mdb";
            this.textBoxPath.TextChanged += new System.EventHandler(this.TextBoxPath_TextChanged);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1433, 707);
            this.tabControl1.TabIndex = 2;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.textBoxCmd);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.textBoxPath);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Location = new System.Drawing.Point(4, 26);
            this.tabPage1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabPage1.Size = new System.Drawing.Size(1425, 677);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "数据库设定";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(878, 48);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "确定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // textBoxCmd
            // 
            this.textBoxCmd.Location = new System.Drawing.Point(12, 147);
            this.textBoxCmd.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.textBoxCmd.Name = "textBoxCmd";
            this.textBoxCmd.Size = new System.Drawing.Size(860, 23);
            this.textBoxCmd.TabIndex = 3;
            this.textBoxCmd.TextChanged += new System.EventHandler(this.TextBoxCmd_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 104);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(196, 17);
            this.label2.TabIndex = 2;
            this.label2.Text = "交互设定(收到指令后进行数据保存)";
            this.label2.Click += new System.EventHandler(this.Label2_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.splitContainer8);
            this.tabPage3.Location = new System.Drawing.Point(4, 26);
            this.tabPage3.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabPage3.Size = new System.Drawing.Size(1425, 677);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "记录查询";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // splitContainer8
            // 
            this.splitContainer8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer8.Location = new System.Drawing.Point(3, 4);
            this.splitContainer8.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.splitContainer8.Name = "splitContainer8";
            this.splitContainer8.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer8.Panel1
            // 
            this.splitContainer8.Panel1.Controls.Add(this.cbx_Model);
            this.splitContainer8.Panel1.Controls.Add(this.label3);
            this.splitContainer8.Panel1.Controls.Add(this.buttonSearch);
            this.splitContainer8.Panel1.Controls.Add(this.textBox_Code);
            this.splitContainer8.Panel1.Controls.Add(this.label15);
            this.splitContainer8.Panel1.Controls.Add(this.dateTimePicker2);
            this.splitContainer8.Panel1.Controls.Add(this.dateTimePicker1);
            this.splitContainer8.Panel1.Controls.Add(this.label14);
            this.splitContainer8.Panel1.Controls.Add(this.label13);
            // 
            // splitContainer8.Panel2
            // 
            this.splitContainer8.Panel2.Controls.Add(this.dataGridViewDynamic1);
            this.splitContainer8.Size = new System.Drawing.Size(1419, 669);
            this.splitContainer8.SplitterDistance = 88;
            this.splitContainer8.SplitterWidth = 7;
            this.splitContainer8.TabIndex = 1;
            // 
            // cbx_Model
            // 
            this.cbx_Model.FormattingEnabled = true;
            this.cbx_Model.Items.AddRange(new object[] {
            "HCBI开关组",
            "HC1AS5加热垫总成30002"});
            this.cbx_Model.Location = new System.Drawing.Point(39, 49);
            this.cbx_Model.Name = "cbx_Model";
            this.cbx_Model.Size = new System.Drawing.Size(334, 25);
            this.cbx_Model.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 52);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(32, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "表名";
            // 
            // buttonSearch
            // 
            this.buttonSearch.Location = new System.Drawing.Point(770, 1);
            this.buttonSearch.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.Size = new System.Drawing.Size(127, 83);
            this.buttonSearch.TabIndex = 6;
            this.buttonSearch.Text = "查询";
            this.buttonSearch.UseVisualStyleBackColor = true;
            this.buttonSearch.Click += new System.EventHandler(this.buttonSearch_Click);
            // 
            // textBox_Code
            // 
            this.textBox_Code.Location = new System.Drawing.Point(432, 51);
            this.textBox_Code.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBox_Code.Name = "textBox_Code";
            this.textBox_Code.Size = new System.Drawing.Size(328, 23);
            this.textBox_Code.TabIndex = 5;
            this.textBox_Code.Text = "010511AS5MA0440001";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(376, 54);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(56, 17);
            this.label15.TabIndex = 4;
            this.label15.Text = "产品条码";
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(432, 3);
            this.dateTimePicker2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(328, 23);
            this.dateTimePicker2.TabIndex = 3;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(93, 7);
            this.dateTimePicker1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(280, 23);
            this.dateTimePicker1.TabIndex = 2;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(397, 7);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(22, 17);
            this.label14.TabIndex = 1;
            this.label14.Text = "->";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(9, 7);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(56, 17);
            this.label13.TabIndex = 0;
            this.label13.Text = "生产时间";
            // 
            // dataGridViewDynamic1
            // 
            this.dataGridViewDynamic1.AllowUserToAddRows = false;
            this.dataGridViewDynamic1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDynamic1.ColumnKey = null;
            this.dataGridViewDynamic1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewDynamic1.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewDynamic1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dataGridViewDynamic1.Name = "dataGridViewDynamic1";
            this.dataGridViewDynamic1.RowHeadersWidth = 51;
            this.dataGridViewDynamic1.RowTemplate.Height = 23;
            this.dataGridViewDynamic1.Size = new System.Drawing.Size(1419, 574);
            this.dataGridViewDynamic1.TabIndex = 0;
            // 
            // FormSave2MdbTool
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1433, 707);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "FormSave2MdbTool";
            this.Text = "FormParameter";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.splitContainer8.Panel1.ResumeLayout(false);
            this.splitContainer8.Panel1.PerformLayout();
            this.splitContainer8.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer8)).EndInit();
            this.splitContainer8.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDynamic1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.SplitContainer splitContainer8;
        private System.Windows.Forms.Button buttonSearch;
        private System.Windows.Forms.TextBox textBox_Code;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private MesDatasCore.DataGridViewDynamic dataGridViewDynamic1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxCmd;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbx_Model;
        private System.Windows.Forms.Button button1;
    }
}