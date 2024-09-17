using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Resources;
using LogConfig;
using HslCommunication.Core;
using HslCommunication;
using UltimateFileSoftwareUpdate.UpdateModel;
using YYUltimateFileClient;

namespace MesDatas
{
    public partial class Fwelcome : Form
    {
        public Form1 form1;
        public string[] Develop = new string[5];
        public string[] DeveloppassWord = new string[5];
        public string[] Admin = new string[5];
        public string[] AdminpassWord = new string[5];
        public string[] Eng = new string[5];
        public string[] EngpassWord = new string[5];
        public string[] Emp = new string[5];
        public string[] EmppassWord = new string[5];
        IniFiles ini = new IniFiles(Application.StartupPath + @"Model.INI");
        public int access;
        public int accesscard;
        public int access_take;
        public int offLine;
        public int checkcard;
        public int MES = 0;
        public bool IsRunningCheckCard = true;
        public FormCheckCard FormCheckCard;
        public Form工单 Form工单 = new Form工单();
        IniFiles ini_user = new IniFiles(Application.StartupPath + @"Userdata.INI");
        public Stopwatch sw = new Stopwatch();
        mdbDatas mdb = null;
        public static string path4 = System.AppDomain.CurrentDomain.BaseDirectory + "SystemDateBase.mdb";
        public static string userFileuRL = "D:\\BYD_Users\\Users_Data.MDB";
        public static SytemSetEntity sytemSet = null;
        public static List<UserInfoEntity> user = null;
        public string loginUserName;
        public enum Language
        {
            ChineseSimplified,//简体中文
            English, //英语
            Thai //泰语
        }
        //当前选择的语言
        private Language CurrentSelectedLanguage = Language.ChineseSimplified;

        public Fwelcome()
        {

            // sw.Start();
            InitializeComponent();
            //sw.Stop();
            //MessageBox.Show(Convert.ToString("1" + ":" + sw.ElapsedMilliseconds) + "\r\n");
            //sw.Restart();
            //this.WindowState = FormWindowState.Maximized;

            //判断数据库文件是否存在
            bool userFile = File.Exists(userFileuRL);
            if (!userFile)
            {
                //创建用户数据库
                addUserMdb(userFileuRL);
            }

            //获取设备名称信息
            GetDeviceName();
            //获取用户信息
            GetUserInfo();
            GetUpdateExe();
        }
        private string UpateExePath = System.AppDomain.CurrentDomain.BaseDirectory + "\\UpdatDebug\\UltimateFileSoftwareUpdate.exe";
        public async void GetUpdateExe()
        {
            await Task.Run(() =>
            {
                try
                {
                    ConiferFile coniferFile = ConiferFile.GetJson();
                    // 启动客户端，此处按照实际项目需求放到了后台线程处理，事实上这种耗时的操作就应该放到后台线程
                    UltimateFileClientFactory ultimateFileClientFactory = new UltimateFileClientFactory();
                    ultimateFileClientFactory.UltimateFileClientStart(coniferFile.IP, coniferFile.Port, coniferFile.Token);
                    OperateResult<GroupFileItem[]> result = ultimateFileClientFactory.integrationFileClient.DownloadPathFileNames(coniferFile.Path, "", "");
                    if (result.IsSuccess)
                    {
                        GroupFileItem[] files = result.Content;
                        // dic.Clear();
                        //int index = 1;
                        foreach (var file in files)
                        {
                            try
                            {
                                TreeNode node = new TreeNode(file.FileName);
                                node.Tag = file;
                                if (!string.IsNullOrEmpty(file.Description))
                                {
                                    int versionA = ConiferFile.CompareVersions(coniferFile.Version, file.Description);
                                    if (versionA < 0)
                                    {
                                        if (MessageBox.Show("新版本：" + file.Description + Environment.NewLine + "是否确认更新？", "新版", MessageBoxButtons.YesNo) == DialogResult.Yes)
                                        {
                                            // 确认更新，启动当前目录的更新程序
                                            try
                                            {
                                                System.Diagnostics.Process.Start(UpateExePath);
                                                System.Threading.Thread.Sleep(20);
                                                //  coniferFile.Version = file.Description;
                                                // MessageBox.Show("有新版本，请更新");
                                                Environment.Exit(0);
                                            }
                                            catch (Exception)
                                            {
                                                MessageBox.Show("更新失败，请手动更新");
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception)
                            {

                            }
                            // dic.Add(index,node);
                        }
                        // treeView1.Nodes[0].Expand();
                    }
                    else
                    {

                    }
                }
                catch { }
            });
        }
        Assembly asm = Assembly.GetExecutingAssembly();
        ResourceManager rm = null;
        public void Fwelcome_Load(object sender, EventArgs e)
        {
            NLogHelperYY.InitConfigNLog();
            tbx_userID.Focus();
            tbx_userID.Select();
            this.label4.Text = sytemSet.DeviceName;


            FontStyle fontStyle = FontStyle.Bold;//设置字体粗细
            float size = 42F;//字体大小
            changeLabelFont(label4, size, fontStyle);//内字体随着字数的增加而自动减小




            ////设置combobox的值
            string language = Properties.Settings.Default.DefaultLanguage;
            if (language == "zh-CN")
            {
                comboBox3.SelectedIndex = 0;
                rm = new ResourceManager("MesDatas.Language_Resources.language_Chinese", asm);
            }
            else if (language == "en-US")
            {
                comboBox3.SelectedIndex = 1;
                rm = new ResourceManager("MesDatas.Language_Resources.language_English", asm);
            }
            else if (language == "th-TH")
            {
                comboBox3.SelectedIndex = 2;
                rm = new ResourceManager("MesDatas.Language_Resources.language_Thai", asm);
            }
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox1.Items.Add(rm.GetString("loginMode"));
            comboBox1.Items.Add(rm.GetString("loginMode1"));
            comboBox2.Items.Add(rm.GetString("comboBox2.Items"));
            comboBox2.Items.Add(rm.GetString("comboBox2.Items1"));
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
        }

        #region Label内字体随着字数的增加而自动减小，Label大小不变
        public Label changeLabelFont(Label label, float size, FontStyle fontStyle)
        {
            Color color = label.ForeColor;
            //FontStyle fontStyle = FontStyle.Bold;
            System.Drawing.FontFamily ff = new System.Drawing.FontFamily(label.Font.Name);
            //float size = 42F;
            string content = label.Text;
            //初始化label状态
            label.Font = new Font(ff, size, fontStyle, GraphicsUnit.Point);
            while (true)
            {
                //获取当前一行能放多少个字======================================================
                //1、获取label宽度
                int labelwidth = label.Width;
                //2、获取当前字体宽度
                Graphics gh = label.CreateGraphics();
                SizeF sf = gh.MeasureString("0", label.Font);
                float fontwidth = sf.Width;
                //3、得到一行放几个字
                int OneRowFontNum = (int)((double)labelwidth / (double)fontwidth);


                //判断当前的Label能放多少列======================================================
                //1、获取当前字体的高度
                float fontheight = sf.Height;
                //2、获取当前label的高度
                int labelheight = label.Height;
                //3、得到当前label能放多少列
                int ColNum = (int)((double)labelheight / (double)fontheight);

                //获取当前字符串需要放多少列======================================================
                var NeedColNum = Math.Ceiling((double)content.Length / (double)OneRowFontNum);

                //如果超出范围，则缩小字体，然后返回再判断一次===================================
                if (ColNum < NeedColNum)
                {
                    size -= 0.25F;
                    label.Font = new Font(ff, size, fontStyle, GraphicsUnit.Point);
                }
                else
                {
                    break;
                }
            }

            return label;
        }
        #endregion

        private void Fwelcome_FormClosed(object sender, FormClosedEventArgs e)
        {
            //LineModel_Write();
            //IsRunning = false;
            //StopTools();
            //dataGridView1.updateData -= ProcCompleteData;
            //SaveParam();
            Environment.Exit(0);
        }

        /// <summary>
        /// 创建Users_Data.MDB
        /// </summary>
        /// <param name="conn"></param>
        private void addUserMdb(string conn)
        {
            mdb = new mdbDatas(conn);
            mdbDatas.CreateAccessDatabase(conn);
            mdbDatas.CreateMDBTable(conn, "Users", new System.Collections.ArrayList(new object[] { "用户名", "用户密码", "用户权限", "厂牌UID", "最后登录时间", "登录次数", "工号" }));
            DataTable dt = new DataTable("Users");
            DataColumn userid = new DataColumn("用户名", typeof(string));
            dt.Columns.Add(userid);
            DataColumn userpwd = new DataColumn("用户密码", typeof(string));
            dt.Columns.Add(userpwd);
            DataColumn authority = new DataColumn("用户权限", typeof(string));
            dt.Columns.Add(authority);
            DataColumn cardId = new DataColumn("厂牌UID", typeof(string));
            dt.Columns.Add(cardId);
            DataColumn lastLoginTime = new DataColumn("最后登录时间", typeof(string));
            dt.Columns.Add(lastLoginTime);
            DataColumn loginNum = new DataColumn("登录次数", typeof(string));
            dt.Columns.Add(loginNum);
            DataColumn JobId = new DataColumn("工号", typeof(string));
            dt.Columns.Add(JobId);
            //DataColumn loginType = new DataColumn("登录方式", typeof(string));
            //dt.Columns.Add(loginType);

            //DataRow dr = dt.NewRow();
            //dt.Rows.Add(dr);

            //dr[0] = "开发者";
            //dr[1] = "dev";
            //dr[2] = "DEV";
            //dr[3] = "";
            //dr[4] = "";
            //dr[5] = "0";
            //dr[6] = "dev";
            //dr[7] = "密码";

            mdb.DatatableToMdb("Users", dt);

            mdb.CloseConnection();
        }

        /// <summary>
        /// 获取设备名称
        /// </summary>
        private void GetDeviceName()
        {
            sytemSet = new SytemSetEntity();
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");

            for (int i = 0; i < table1.Rows.Count; i++)
            {
                for (int j = 0; j < table1.Columns.Count; j++)
                {

                    sytemSet.DeviceName = table1.Rows[i]["DeviceName"].ToString();
                    sytemSet.rfidDeviceCode = table1.Rows[i]["rfidCode"].ToString();
                    sytemSet.rfidPort = table1.Rows[i]["rfidProt"].ToString();

                }
            }
            mdb.CloseConnection();
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
                userInfo.userName = table1.Rows[i]["用户名"].ToString();
                user.Add(userInfo);
            }
            mdb.CloseConnection();
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            access = 0;
            access_take = 0;//在线
            if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex == 0)
            {
                offLine = 0;
                checkcard = 0;
            }
            else if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex == 1)
            {
                offLine = 0;
                checkcard = 1;
            }//离线
            else if (comboBox1.SelectedIndex == 1 && comboBox2.SelectedIndex == 0)
            {
                offLine = 1;
                checkcard = 0;
            }
            else if (comboBox1.SelectedIndex == 1 && comboBox2.SelectedIndex == 1)
            {
                offLine = 1;
                checkcard = 1;
            }

            else if (comboBox1.Text == "请选择登录模式" || comboBox1.Text == "")
            {
                offLine = 2;
                checkcard = 2;
                MessageBox.Show(rm.GetString("ModeTip"));
            }
            else if (comboBox2.Text == "请选择登录类型" || comboBox2.Text == "")
            {
                offLine = 2;
                checkcard = 2;
                MessageBox.Show("请选择登录类型！");
            }

            if (checkcard == 0)
            {
                if (tbx_userID.Text != "" && tbx_Password.Text != "")
                {
                    // 根据JSON配置文件动态配置dev权限的账号密码
                    var devUsername = ConfigManager.GetConfigValue("DevUsername");
                    var devPassword = ConfigManager.GetConfigValue("DevPassword");

                    if (tbx_userID.Text == devUsername && tbx_Password.Text == devPassword)
                    {
                        access = 4;
                        access_take = 1;
                    }
                    else
                    {

                        List<UserInfoEntity> list = user.Where(x => x.Uuser == tbx_userID.Text && x.Upwd == tbx_Password.Text).ToList();

                        if (list.Count > 0)
                        {

                            foreach (var v in list)
                            {
                                if (tbx_userID.Text == v.Uuser && tbx_Password.Text == v.Upwd)//开发
                                {

                                    //判断是否有离线登录权限
                                    if ((v.Utype != "ADM" && v.Utype != "PE" && v.Utype != "QE") && comboBox1.SelectedIndex == 1)
                                    {
                                        MessageBox.Show(rm.GetString("offlineTip"));
                                        return;
                                    }

                                    loginUserName = v.userName;

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
                                }

                            }
                        }
                        else
                        {
                            MessageBox.Show(rm.GetString("checkIdorpwd"));
                            access_take = 0;
                            return;
                        }
                    }

                    //else if (tbx_userID.Text == user.Uuser && tbx_Password.Text == user.Upwd)//管理员
                    //{
                    //    Access = 2;
                    //    Access_take = 1;
                    //}
                    //else if (tbx_userID.Text == user.Uuser && tbx_Password.Text == user.Upwd)//操作员
                    //{
                    //    Access = 1; 
                    //    Access_take = 1;
                    //}


                }
                else
                {
                    access = 0;
                    access_take = 0;
                    MessageBox.Show(rm.GetString("inputTip"));
                }
                if (access > 0)
                {
                    this.Hide();//隐藏窗体
                    form1 = new Form1();
                    form1.SetMES(1);
                    form1.SetoffLineType(offLine);
                    form1.Setcheckcard(checkcard);
                    form1.Setaccess_take(access_take);
                    form1.Setaccess(access);
                    form1.SetloginName(loginUserName);
                    form1.SetloginUser(tbx_userID.Text);
                    form1.SetloginPwd(tbx_Password.Text);
                    form1.Show();
                }

            }
            //else if (CheckCard == 1)
            //{
            //    Hide();
            //    FormCheckCard.Show();
            //}
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        /// <summary>
        /// 回车键事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void button1_KeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Button1_Click(null, null);
            }
        }

        /// <summary>
        /// 选中刷卡事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 1)
            {

                if (comboBox1.SelectedIndex == 1 && comboBox2.SelectedIndex == 0)
                {
                    offLine = 1;
                    checkcard = 0;
                }
                else if (comboBox1.SelectedIndex == 1 && comboBox2.SelectedIndex == 1)
                {
                    offLine = 1;
                    checkcard = 1;
                }
                else if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex == 0)
                {
                    offLine = 0;
                    checkcard = 0;
                }
                else if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex == 1)
                {
                    offLine = 0;
                    checkcard = 1;
                }

                else if (comboBox1.Text == "请选择登录模式" || comboBox1.Text == "")
                {
                    offLine = 2;
                    checkcard = 2;
                    MessageBox.Show(rm.GetString("ModeTip"));
                }
                else if (comboBox2.Text == "请选择登录类型" || comboBox2.Text == "")
                {
                    offLine = 2;
                    checkcard = 2;
                    MessageBox.Show("请选择登录类型！");
                }

                this.Hide();//隐藏窗体
                FormCheckCard = new FormCheckCard();
                FormCheckCard.SetMES(1);
                FormCheckCard.SetoffLineType(offLine);
                FormCheckCard.Setcheckcard(checkcard);
                FormCheckCard.Setaccess_take(access_take);
                FormCheckCard.Setaccess(access);
                FormCheckCard.SetrfidDeviceCode(sytemSet.rfidDeviceCode);
                FormCheckCard.SetrfidPort(sytemSet.rfidPort);
                FormCheckCard.Show();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Enabled = false;
            if (comboBox3.SelectedIndex == 0)
            {
                //修改默认语言
                MultiLanguage.SetDefaultLanguage("zh-CN");
                this.CurrentSelectedLanguage = Language.ChineseSimplified;
                rm = new ResourceManager("MesDatas.Language_Resources.language_Chinese", asm);
                //对所有打开的窗口重新加载语言
                foreach (Form form in Application.OpenForms)
                {
                    LoadAll(form);
                }

            }
            else if (comboBox3.SelectedIndex == 1)
            {
                //修改默认语言
                MultiLanguage.SetDefaultLanguage("en-US");
                this.CurrentSelectedLanguage = Language.English;
                rm = new ResourceManager("MesDatas.Language_Resources.language_English", asm);
                //对所有打开的窗口重新加载语言
                foreach (Form form in Application.OpenForms)
                {
                    LoadAll(form);
                }

            }
            else if (comboBox3.SelectedIndex == 2)
            {
                //修改默认语言
                MultiLanguage.SetDefaultLanguage("th-TH");
                this.CurrentSelectedLanguage = Language.Thai;
                rm = new ResourceManager("MesDatas.Language_Resources.language_Thai", asm);
                //对所有打开的窗口重新加载语言
                foreach (Form form in Application.OpenForms)
                {
                    LoadAll(form);
                }
            }
            comboBox3.Enabled = true;
        }

        private void LoadAll(Form form)
        {
            if (form.Name == "Fwelcome")
            {
                MultiLanguage.LoadLanguage(form, typeof(Fwelcome));
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();

                comboBox1.Items.Add(rm.GetString("loginMode"));
                comboBox1.Items.Add(rm.GetString("loginMode1"));
                comboBox2.Items.Add(rm.GetString("comboBox2.Items"));
                comboBox2.Items.Add(rm.GetString("comboBox2.Items1"));
                comboBox1.SelectedIndex = 0;
                comboBox2.SelectedIndex = 0;
                this.label4.Text = sytemSet.DeviceName;
            }
            else if (form.Name == "Form1")
            {
                MultiLanguage.LoadLanguage(form, typeof(Form1));
            }
            else if (form.Name == "Form2")
            {
                MultiLanguage.LoadLanguage(form, typeof(Form2));
            }

        }


    }
}
