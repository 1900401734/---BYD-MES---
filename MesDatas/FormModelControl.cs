using HslCommunication.Profinet.Omron;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HslCommunication;

namespace MesDatas
{
    public partial class FormModelControl : Form
    {
        public string[] Parameter_Model = new string[25000];
        public string PLClink;
        IniFiles ini = new IniFiles(Application.StartupPath + @"Model.INI");
        public FormModelControl()
        {
            InitializeComponent();
          
        }
        private void FormModelControl_Load(object sender, EventArgs e)
        {
           LogMsg_Model();
        }
        private void FormModelControl_FormClosed(object sender, FormClosedEventArgs e)
        {   if (PLClink == "1")
            { omronFinsNet.ConnectClose(); }
            else
            { }
            
        }

        private void LogMsg_Model()
        {
            //this.Invoke(new Action(() =>
            //{

            //}));

            if (richTextBox1.TextLength > 50000)
            {
                richTextBox1.Clear();
            }
            for (int i = 1; i < 100; i++)//ctrl+H
            {

                Parameter_Model[i] = ini.IniReadValue(i.ToString(), "型号" );
                Parameter_Model[i + 1000] = ini.IniReadValue(i.ToString(), "高度上限");
                Parameter_Model[i + 1500] = ini.IniReadValue(i.ToString(), "高度下限");
                Parameter_Model[i + 17000] = ini.IniReadValue(i.ToString(), "PLC配方");
                Parameter_Model[i + 17500] = ini.IniReadValue(i.ToString(), "文件版本");
                Parameter_Model[i + 17800] = ini.IniReadValue(i.ToString(), "软件版本");
                Parameter_Model[i + 18000] = ini.IniReadValue(i.ToString(), "产品编码");
                Parameter_Model[i + 18500] = ini.IniReadValue(i.ToString(), "产品名称");
                Parameter_Model[i + 19000] = ini.IniReadValue(i.ToString(), "产品代码");
                Parameter_Model[i + 19500] = ini.IniReadValue(i.ToString(), "工装代码");
                richTextBox1.AppendText(i.ToString() + ":" + Parameter_Model[i] + "\r\n"+ "高度上限:" + Parameter_Model[i + 1000] + "\r\n" + "高度下限:" + Parameter_Model[i + 1500] + "\r\n" 
                    +  "PLC配方:" + Parameter_Model[i + 17000] + "\r\n" +
                    "文件版本:" + Parameter_Model[i + 17500] + "\r\n" + "软件版本:" + Parameter_Model[i + 17800] + "\r\n" + "产品编码:" + Parameter_Model[i + 18000] + "\r\n" + "产品名称:" + Parameter_Model[i + 18500] + "\r\n" + 
                    "产品代码:" + Parameter_Model[i + 19000] + "\r\n" + "工装代码:" + Parameter_Model[i + 19500] + "\r\n");
                richTextBox1.Select(0, 0);
                richTextBox1.ScrollToCaret();
            }

        }
        private OmronFinsNet omronFinsNet;
        private void FinsTcp_Connect()
        {
            try
            {
                omronFinsNet = new OmronFinsNet("192.168.0.10", Convert.ToInt16("9600"));
                omronFinsNet.SA1 = 34;    // PC网络号，PC的IP地址的最后一个数
                omronFinsNet.DA1 = 0x01;  // PLC网络号，PLC的IP地址的最后一个数
                omronFinsNet.DA2 = 0x00; // PLC单元号，通常为0
                OperateResult connect = omronFinsNet.ConnectServer();
                if (connect.IsSuccess)
                {
                    MessageBox.Show("PLC连接成功");
                    PLClink = "1";
                }
                else
                {
                    MessageBox.Show("PLC连接失败,请重启软件！");
                    PLClink = "0";
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
        private void Button1_Click(object sender, EventArgs e)
        {

            //for (int i = 0; i < 100; i++)
            //{
                //ini.IniWriteValue(i.ToString(), "型号", "NO" + i.ToString());
                //ini.IniWriteValue(i.ToString(), "产品编码", "1" + i.ToString());
                //ini.IniWriteValue(i.ToString(), "产品名称", "2" + i.ToString());
                //ini.IniWriteValue(i.ToString(), "产品代码", "3" + i.ToString());
                //ini.IniWriteValue(i.ToString(), "工装代码", "4" + i.ToString());
                //ini.IniWriteValue(i.ToString(), "高度上限", "5");
                //ini.IniWriteValue(i.ToString(), "高度下限", "6");
                //ini.IniWriteValue(i.ToString(), "PLC配方", "1");
                //ini.IniWriteValue(i.ToString(), "文件版本", "20220811");
                //ini.IniWriteValue(i.ToString(), "软件版本", "20220811");
            //    ini.IniWriteValue(i.ToString(), "条码规则", "030510XXXX");
            //}
            if (tbx_Number.Text != "" && tbx_Modelname.Text != "")
            {
                ini.IniWriteValue(tbx_Number.Text, "型号", tbx_Modelname.Text);
                ini.IniWriteValue(tbx_Number.Text, cbx_spec.Text, tbx_setvalue.Text);
                MessageBox.Show("参数保存成功！");
            }
            else
            { MessageBox.Show("保存失败，请检查编号、型号！"); }


        }

       
    }
}
