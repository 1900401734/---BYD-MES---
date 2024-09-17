using MesDatasCore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ToolSample
{
    public partial class FormToolSample : Form
    {
        public FormToolSample()
        {
            InitializeComponent();

        }

        private cToolBase cTool = null;
        public void SetTool(cToolBase Tool)
        {
            cTool = Tool;
        }

        private string ReadString(string key)
        {
            return cTool.ReadString(cTool.sName, key);
        }

        private void WriteString(string key, string value)
        {
            cTool.WriteString(cTool.sName, key, value);
        }

        public void LoadParameter()
        {
            // ReadString
        }


        public void SaveParameter()
        {
            // WriteString
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("button1_Click");
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
