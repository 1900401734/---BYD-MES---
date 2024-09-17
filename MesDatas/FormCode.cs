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
    public partial class FormCode : Form
    {
        IniFiles ini = new IniFiles(Application.StartupPath + @"Model.INI");
        string[] Parameter_Code = new string[25000];
        
        public FormCode()
        {
            InitializeComponent();
        }
        private void FormCode_Load(object sender, EventArgs e)
        {
            Model_Read_Other();
        }

        private void Model_Read_Other()
        {

            for (int i = 1; i < 100; i++)
            {

                Parameter_Code[i] = ini.IniReadValue(i.ToString(), "型号");
                comboBox1.Items.Add( Parameter_Code[i]);
                Parameter_Code[i+1000] = ini.IniReadValue(i.ToString(), "条码规则");

            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (tbx_Num.Text != "" && textBox1.Text != "")
            {
                ini.IniWriteValue(tbx_Num.Text, "条码规则", textBox1.Text);
                this.Close();
            }
            else
            { MessageBox.Show("保存失败，请检查编号、型号！"); }
        }


        private void Button1_Click(object sender, EventArgs e)
        { DateTime times_Month = DateTime.Now;
         string times_Month_string0 = times_Month.Year.ToString();
         string times_Month_string1 = times_Month.Month.ToString();
         string times_Month_string2 = times_Month.Day.ToString();
        Invoke(new Action(() =>
            {
                for (int i = 1; i < 100; i++)
                {
                    Parameter_Code[i] = ini.IniReadValue(i.ToString(), "型号");
                    Parameter_Code[i+1000] = ini.IniReadValue(i.ToString(), "条码规则");
                }

                for (int i = 1; i < 100; i++)
            {

                if (comboBox1.Text == Parameter_Code[i])
                {
                    Parameter_Code[0] = Parameter_Code[i];
                    Parameter_Code[1000] = Parameter_Code[i + 1000];
                    ini.IniWriteValue("0", "型号", Parameter_Code[0]);
                    ini.IniWriteValue("0", "条码规则", Parameter_Code[1000]);
                    comboBox1.Text = Parameter_Code[0];
                    textBox1.Text = Parameter_Code[1000];
                    tbx_Num.Text = i.ToString();
                }
            }
            }));
        }
    }
}
