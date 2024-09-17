using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MesDatas
{
    public partial class Form工单 : Form
    {
        IniFiles ini = new IniFiles(Application.StartupPath + @"Model.INI");
        public Form工单()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            ini.IniWriteValue("工单", "工单号", textBox1.Text);
            this.Close();
        }
    }
}
