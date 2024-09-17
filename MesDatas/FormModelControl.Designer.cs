namespace MesDatas
{
    partial class FormModelControl
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
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tbx_Number = new System.Windows.Forms.TextBox();
            this.编号 = new System.Windows.Forms.Label();
            this.型号名 = new System.Windows.Forms.Label();
            this.tbx_Modelname = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbx_setvalue = new System.Windows.Forms.TextBox();
            this.cbx_spec = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.richTextBox1.Location = new System.Drawing.Point(7, 6);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(571, 973);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("微软雅黑", 26.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(599, 411);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(252, 96);
            this.button1.TabIndex = 1;
            this.button1.Text = "确认更改";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // tbx_Number
            // 
            this.tbx_Number.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbx_Number.Location = new System.Drawing.Point(594, 59);
            this.tbx_Number.Name = "tbx_Number";
            this.tbx_Number.Size = new System.Drawing.Size(257, 26);
            this.tbx_Number.TabIndex = 2;
            // 
            // 编号
            // 
            this.编号.AutoSize = true;
            this.编号.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.编号.Location = new System.Drawing.Point(594, 29);
            this.编号.Name = "编号";
            this.编号.Size = new System.Drawing.Size(52, 27);
            this.编号.TabIndex = 3;
            this.编号.Text = "编号";
            // 
            // 型号名
            // 
            this.型号名.AutoSize = true;
            this.型号名.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.型号名.Location = new System.Drawing.Point(594, 121);
            this.型号名.Name = "型号名";
            this.型号名.Size = new System.Drawing.Size(72, 27);
            this.型号名.TabIndex = 5;
            this.型号名.Text = "型号名";
            // 
            // tbx_Modelname
            // 
            this.tbx_Modelname.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbx_Modelname.Location = new System.Drawing.Point(594, 154);
            this.tbx_Modelname.Name = "tbx_Modelname";
            this.tbx_Modelname.Size = new System.Drawing.Size(257, 26);
            this.tbx_Modelname.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(594, 213);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 27);
            this.label1.TabIndex = 7;
            this.label1.Text = "规格项";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(594, 305);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 27);
            this.label2.TabIndex = 9;
            this.label2.Text = "设定值";
            // 
            // tbx_setvalue
            // 
            this.tbx_setvalue.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbx_setvalue.Location = new System.Drawing.Point(594, 345);
            this.tbx_setvalue.Name = "tbx_setvalue";
            this.tbx_setvalue.Size = new System.Drawing.Size(257, 26);
            this.tbx_setvalue.TabIndex = 8;
            // 
            // cbx_spec
            // 
            this.cbx_spec.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold);
            this.cbx_spec.FormattingEnabled = true;
            this.cbx_spec.Items.AddRange(new object[] {
            "高度上限",
            "高度下限",
            "PLC配方",
            "产品编码",
            "产品名称",
            "产品代码",
            "工装代码",
            "文件版本",
            "软件版本"});
            this.cbx_spec.Location = new System.Drawing.Point(594, 249);
            this.cbx_spec.Name = "cbx_spec";
            this.cbx_spec.Size = new System.Drawing.Size(257, 27);
            this.cbx_spec.TabIndex = 10;
            // 
            // FormModelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(863, 985);
            this.Controls.Add(this.cbx_spec);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tbx_setvalue);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.型号名);
            this.Controls.Add(this.tbx_Modelname);
            this.Controls.Add(this.编号);
            this.Controls.Add(this.tbx_Number);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.richTextBox1);
            this.Name = "FormModelControl";
            this.Text = "FormModelControl";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormModelControl_FormClosed);
            this.Load += new System.EventHandler(this.FormModelControl_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox tbx_Number;
        private System.Windows.Forms.Label 型号名;
        private System.Windows.Forms.TextBox tbx_Modelname;
        private System.Windows.Forms.Label 编号;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbx_setvalue;
        private System.Windows.Forms.ComboBox cbx_spec;
    }
}