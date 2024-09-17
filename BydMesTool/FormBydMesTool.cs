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
using 工艺部信息化组;

namespace BydMesTool
{
    public partial class FormBydMesTool : Form
    {
        public FormBydMesTool()
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

        private string ip => textBox_ip.Text;
        private string port => textBox_port.Text;
        private string timeout => textBox_timeout.Text;
        private string nccode => textBox_nccode.Text;
        private string operation => textBox_opration.Text;
        private string user => textBox_user.Text;
        private string password => textBox_password.Text;
        private string url => textBox_url.Text;
        private string site => textBox_site.Text;
        private string resource => textBox_resource.Text;

        public void LoadParameter()
        {
            // mes config
            textBox_ip.Text = ReadString("IP");
            textBox_port.Text = ReadString("Port");
            textBox_timeout.Text = ReadString("Timeout");
            textBox_nccode.Text = ReadString("NcCode");
            textBox_opration.Text = ReadString("Opration");
            textBox_password.Text = ReadString("Password");
            textBox_resource.Text = ReadString("Resource");
            textBox_site.Text = ReadString("Site");
            textBox_url.Text = ReadString("Url");
            textBox_user.Text = ReadString("User");

            //mes cmd
            UserCheckCmd.Text = ReadString("UserCheckCmd");
            UserCheckPass.Text = ReadString("UserCheckPass");
            UserCheckFail.Text = ReadString("UserCheckFail");

            CodeCheckCmd.Text = ReadString("CodeCheckCmd");
            CodeCheckPass.Text = ReadString("CodeCheckPass");

            CodeSendCmd.Text = ReadString("CodeSendCmd");


        }


        public void SaveParameter()
        {
            // mes config
            WriteString("IP", textBox_ip.Text);
            WriteString("Port", textBox_port.Text);
            WriteString("Timeout", textBox_timeout.Text);
            WriteString("NcCode", textBox_nccode.Text);
            WriteString("Opration", textBox_opration.Text);
            WriteString("Password", textBox_password.Text);
            WriteString("Resource", textBox_resource.Text);
            WriteString("Site", textBox_site.Text);
            WriteString("Url", textBox_url.Text);
            WriteString("User", textBox_user.Text);

            //mes cmd            
            WriteString("UserCheckCmd", UserCheckCmd.Text);
            WriteString("UserCheckPass", UserCheckPass.Text);
            WriteString("UserCheckFail", UserCheckFail.Text);

            WriteString("CodeCheckCmd", CodeCheckCmd.Text);
            WriteString("CodeCheckPass", CodeCheckPass.Text);

            WriteString("CodeSendCmd", CodeSendCmd.Text);


        }

        public void Config_Mes(string ip, string port, string timeout,
           string url, string site, string user, string password, string resource, string operation, string ncCode)
        {
            工艺部信息化组.CONFIG.IP = ip;
            工艺部信息化组.CONFIG.PORT = port;
            工艺部信息化组.CONFIG.TimeOut = int.Parse(timeout);
            工艺部信息化组.CONFIG.URL = url;
            工艺部信息化组.CONFIG.Site = site;
            工艺部信息化组.CONFIG.UserName = user;
            工艺部信息化组.CONFIG.Password = password;
            工艺部信息化组.CONFIG.Resource = resource;
            工艺部信息化组.CONFIG.Operation = operation;
            工艺部信息化组.CONFIG.NcCode = ncCode;
        }

        public string MesUserCheck { get { return UserCheckCmd.Text; } }
        public string MesBarCodeCheck { get { return CodeCheckCmd.Text; } }
        public string MesBarCodeSend { get { return CodeSendCmd.Text; } }

        public void UsersVarify(out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
            BydMesCom.用户验证(out 验证结果, out MES反馈, out XMLOUT);
        }

        public void BarCodeVarify(string 产品条码, out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
            BydMesCom.条码验证(产品条码, out 验证结果, out MES反馈, out XMLOUT);
        }

        public void UpDateToMes(bool 测试结果, string 产品条码, string 文件版本, string 软件版本, string 测试项, out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
            BydMesCom.条码上传(测试结果, 产品条码, 文件版本, 软件版本, 测试项, out 验证结果, out MES反馈, out XMLOUT);
        }

        private void userClick(object sender, EventArgs e)
        {
            Config_Mes(ip, port, timeout, url, site, user, password, resource, operation, nccode);
            bool 验证结果;
            string MES反馈;
            string XMLOUT;
            UsersVarify(out 验证结果, out MES反馈, out XMLOUT);

        }

        private void Barcode_Click(object sender, EventArgs e)
        {
            string 产品条码 = textBox1.Text; 
            bool 验证结果;
            string MES反馈;
            string XMLOUT;
            BarCodeVarify(产品条码, out 验证结果, out MES反馈, out XMLOUT);
        }

        private void CodeVarify_Click(object sender, EventArgs e)
        {
            bool 测试结果 = false; string 产品条码 = textBox1.Text;  string 文件版本 = "00"; string 软件版本 = "01"; string 测试项 = richTextBox1.Text; bool 验证结果; string MES反馈; string XMLOUT;
            UpDateToMes(测试结果, 产品条码, 文件版本, 软件版本, 测试项, out 验证结果, out MES反馈, out XMLOUT);
        }

        private void 用户登录_Click(object sender, EventArgs e)
        {

        }

        private void 条码验证_Click(object sender, EventArgs e)
        {

        }

        private void 数据上传_Click(object sender, EventArgs e)
        {

        }
    }
}
