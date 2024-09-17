using System;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using HFrfid;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.ComponentModel;
using System.Reflection;
using System.Resources;

namespace MesDatas
{
    public partial class FormCheckCard : Form
    {

        RfidReader Reader = new RfidReader();
        public Form1 form1;
        public Fwelcome fwelcome;
        private int MES { get; set; }
        public void SetMES(int strText)
        {
            MES = strText;
        }
        /// <summary>
        /// 0=联机  1=单机
        /// </summary>
        private int offLineType { get; set; }
        public void SetoffLineType(int strText)
        {
            offLineType = strText;
        }
        private int checkcard { get; set; }
        public void Setcheckcard(int strText)
        {
            checkcard = strText;
        }
        private int access_take { get; set; }
        public void Setaccess_take(int strText)
        {
            access_take = strText;
        }
        /// <summary>
        /// 登入用户类型
        /// </summary>
        private int access { get; set; }
        public void Setaccess(int strText)
        {
            access = strText;
        }
        private string rfidDeviceCode { get; set; }
        public void SetrfidDeviceCode(string strText)
        {
            rfidDeviceCode = strText;
        }
        private string rfidPort { get; set; }
        public void SetrfidPort(string strText)
        {
            rfidPort = strText;
        }

        public FormCheckCard()
        {
            InitializeComponent();
        }

        public static List<UserInfoEntity> user = null;
        mdbDatas mdb = null;
        public static string path4 = System.AppDomain.CurrentDomain.BaseDirectory + "SystemDateBase.mdb";
        public static string userFileuRL = "D:\\BYD_Users\\Users_Data.MDB";

        Task taskProcess_UserID = null;
        bool IsRunningplc_UserID = true;

        Assembly asm = Assembly.GetExecutingAssembly();
        ResourceManager resources = null;

        public void FormCheckCard_Load(object sender, EventArgs e)
        {

            //读取用户
            GetUserInfo();

            string language = Properties.Settings.Default.DefaultLanguage;
            if (language == "zh-CN")
            {
                resources = new ResourceManager("MesDatas.Language_Resources.language_Chinese", asm);
            }
            else if (language == "en-US")
            {
                resources = new ResourceManager("MesDatas.Language_Resources.language_English", asm);
            }
            else if (language == "th-TH")
            {
                resources = new ResourceManager("MesDatas.Language_Resources.language_Thai", asm);
            }

            try
            {
                if (rfidDeviceCode.Length < 1 || rfidPort.Length < 1)
                {
                    MessageBox.Show(resources.GetString("posType"));
                    label2.Text = resources.GetString("posType");
                    return;
                }
                else
                {
                    bool flg = Reader.Connect(rfidPort, 9600);
                    if (flg == true)
                    {
                        IsRunningplc_UserID = true;
                        label2.Text = resources.GetString("portOK");
                        taskProcess_UserID = new Task(Process_UserID);
                        taskProcess_UserID.Start();
                    }
                    else
                    {
                        label2.Text = resources.GetString("portNG");
                        return;
                    }
                }
            }
            catch
            {
                label2.Text = resources.GetString("portType");
                return;
            }
        }

        /// <summary>
        /// 获取账号信息
        /// </summary>
        private void GetUserInfo()
        {
            user = new List<UserInfoEntity>();
            mdb = new mdbDatas(userFileuRL);
            DataTable table1 = mdb.Find("select * from Users");

            for (int i = 0; i < table1.Rows.Count; i++)
            {
                UserInfoEntity userInfo = new UserInfoEntity();
                userInfo.Uuser = table1.Rows[i]["工号"].ToString();
                userInfo.Upwd = table1.Rows[i]["用户密码"].ToString();
                userInfo.Utype = table1.Rows[i]["用户权限"].ToString();
                //userInfo.loginType = table1.Rows[i]["登录方式"].ToString();
                userInfo.userName = table1.Rows[i]["用户名"].ToString();
                userInfo.cardID = table1.Rows[i]["厂牌UID"].ToString();
                user.Add(userInfo);
            }
            mdb.CloseConnection();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IsRunningplc_UserID = false;
            Reader.DisConnect();
            this.Hide();//隐藏窗体
            fwelcome = new Fwelcome();
            fwelcome.Show();
        }

        private void Process_UserID()
        {
            while (IsRunningplc_UserID)
            {
                this.Invoke(new Action(() =>
                {
                    timer1_Tick(null, null);

                }));

                Thread.Sleep(500);
                Application.DoEvents();
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int status;
            byte[] type = new byte[2];
            byte[] id = new byte[4];

            Reader.Cmd = Cmd.M1_ReadId;//读卡号命令
            Reader.Addr = Convert.ToByte(rfidDeviceCode, 16);//读写器地址,设备号
            Reader.Beep = Beep.On;

            status = Reader.M1_Operation();
            if (status == 0)//读卡成功
            {
                for (int i = 0; i < 2; i++)//获取2字节卡类型
                {
                    type[i] = Reader.RxBuffer[i];
                }
                for (int i = 0; i < 04; i++)//获取4字节卡号
                {
                    id[i] = Reader.RxBuffer[i + 2];
                }
                string cardType = byteToHexStr(type, 2);
                tbx1.Text = cardType;
                string userid = byteToHexStr(id, 4);
                tbx2.Text = userid;
                label2.Text = resources.GetString("readCardOK");

                //根据userid 读取账号密码
                List<UserInfoEntity> list = user.Where(x => x.cardID == userid).ToList();
                if (list.Count > 0)
                {
                    if (list.Count > 1)
                    {
                        MessageBox.Show(resources.GetString("serchType"));
                        return;
                    }

                    foreach (var v in list)
                    {

                        //判断是否有离线登录权限
                        if ((v.Utype != "ADM" && v.Utype != "PE" && v.Utype != "QE") && offLineType == 1)
                        {
                            MessageBox.Show(resources.GetString("offlineTip"));
                            return;
                        }
                        if (v.Uuser.Length > 0 && v.Upwd.Length > 0)//开发
                        {
                            if (v.Utype == "ADM")
                            {
                                access = 3;
                            }
                            else if (v.Utype == "PE")
                            {
                                access = 2;
                            }
                            else if (v.Utype == "DEV")
                            {
                                access = 4;
                            }
                            else if (v.Utype == "QE")
                            {
                                access = 5;
                            }
                            else if (v.Utype == "ME")
                            {
                                access = 6;
                            }
                            else if (v.Utype == "OP")
                            {
                                access = 1;
                            }

                            access_take = 1;

                            this.Hide();//隐藏窗体
                            form1 = new Form1();
                            form1.SetMES(1);
                            form1.SetoffLineType(offLineType);
                            form1.Setcheckcard(checkcard);
                            form1.Setaccess_take(access_take);
                            form1.Setaccess(access);
                            form1.SetloginName(v.userName);
                            form1.SetloginUser(v.Uuser);
                            form1.SetloginPwd(v.Upwd);
                            form1.Show();


                            IsRunningplc_UserID = false;
                            Reader.DisConnect();
                        }

                    }
                }
                else
                {
                    label2.Text = resources.GetString("loginType");
                    return;
                }
            }
            else
            {
                label2.Text = ("状态显示：错误码：" + status.ToString());
            }
        }

        public string byteToHexStr(byte[] bytes, int len)  //数组转十六进制字符
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < len; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }
    }
}
