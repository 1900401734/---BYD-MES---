using HFrfid;
using HslCommunication;
using HslCommunication.ModBus;
using HslCommunication.Profinet.Keyence;
using INIFile;
using MesDatas.Entity;
using MesDatas.SqlConverter;
using MesDatas.Utiey;
using MesDatasCore;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UltimateFileSoftwareUpdate.UpdateModel;
using 工艺部信息化组;

namespace MesDatas
{
    public partial class Form1 : Form, IMesDatasBase
    {
        public static int iOperCount = 0;
        public static System.Timers.Timer timer;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000; // Turn on WS_EX_COMPOSITED
                return cp;
            }
        }

        private int MES { get; set; }
        public void SetMES(int strText)
        {
            MES = strText;
        }

        /// <summary>
        /// 登录模式(联机或单机)
        /// </summary>
        public string LoginMode { get; private set; }
        /// <summary>
        /// 0=联机  1=单机
        /// </summary>
        private int OffLineType { get; set; }
        public void SetoffLineType(int strText)
        {
            OffLineType = strText;
            LoginMode = OffLineType == 0 ? "联机" : "单机";
        }

        /// <summary>
        /// 登录方式(刷卡或密码)
        /// </summary>
        public string LoginMethod { get; private set; }
        private int CheckCard { get; set; }
        public void Setcheckcard(int strText)
        {
            CheckCard = strText;
            LoginMethod = CheckCard == 0 ? "密码登录" : "刷卡登录";
        }

        private int Access_take { get; set; }
        public void Setaccess_take(int strText)
        {
            Access_take = strText;
        }

        /// <summary>
        /// 用户权限
        /// </summary>
        private int Access { get; set; }
        public void Setaccess(int strText)
        {
            Access = strText;
            GetAccessName(Access);
        }
        /// <summary>
        /// 权限名
        /// </summary>
        private string AccessName { get; set; }
        public void GetAccessName(int access)
        {
            switch (access)
            {
                case 1:
                    AccessName = "操作员(OP)";
                    break;
                case 2:
                    AccessName = "工艺工程师(PE)";
                    break;
                case 3:
                    AccessName = "管理员(ADM)";
                    break;
                case 4:
                    AccessName = "开发者(DEV)";
                    break;
                case 5:
                    AccessName = "品质工程师(QE)";
                    break;
                case 6:
                    AccessName = "设备工程师(ME)";
                    break;
                default:
                    AccessName = string.Empty;
                    break;
            }
        }

        /// <summary>
        /// 登录用户(工号)
        /// </summary>
        private string LoginUser { get; set; }
        public void SetloginUser(string strText)
        {
            LoginUser = strText;
        }

        /// <summary>
        /// 登录名称（姓名）
        /// </summary>
        private string LoginName { get; set; }
        public void SetloginName(string strText)
        {
            if (strText == null)
            {
                LoginName = "开发者";
            }
            else
            {
                LoginName = strText;
            }
        }

        /// <summary>
        /// 登录密码
        /// </summary>s
        private string LoginPwd { get; set; }
        public void SetloginPwd(string strText)
        {
            LoginPwd = strText;
        }

        //public Fwelcome Fwelcome = new Fwelcome();
        public FormCheckCard FormCheckCard = new FormCheckCard();
        public FormModelControl FormModelControl = new FormModelControl();
        public FormUser FormUser = new FormUser();
        public FormEngineer FormEngineer = new FormEngineer();
        public FormSuperUser FormSuperUser = new FormSuperUser();
        public FormCode FormCode = new FormCode();
        public Form工单 Form工单 = new Form工单();
        public FormCardLogin FormCardLogin = new FormCardLogin();
        private ModbusTcpNet busTcpClient = null;
        public string[] Parameter_txt = new string[10000];
        public string[] Parameter_txt1 = new string[100];
        List<string> list = null;
        List<string> beatlist = null;//节拍
        List<string> maxlist = null;//上限
        List<string> minlist = null;//下限
        List<string> resultlist = null;
        List<string> stationlist = null;
        List<string> workstNamelist = null;
        //List<string> vulpalist = null;//易损件使用数目
        mdbDatas mdb = null;
        public string 工单号;
        public bool 产品结果;
        public int Num = 1;
        public string barcodeInfo = null;
        int[] Count = new int[10000];
        string[] Parameter = new string[10000];
        string[] Value = new string[10000];
        public string[] userdata = new string[1000];
        string[] Parameter_Model = new string[25000];

        public CSVDeal myCSVDeal = new CSVDeal();
        public Stopwatch sw = new Stopwatch();
        public string Runtime = DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日" + DateTime.Now.Hour.ToString() + "：" + DateTime.Now.Minute.ToString() + "：" + DateTime.Now.Second.ToString();
        public enum Language
        {
            ChineseSimplified,//简体中文
            English, //英语
            Thai //泰语
        }

        Assembly asm = Assembly.GetExecutingAssembly();
        ResourceManager resources = null;

        public Form1()
        {
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

            //Fwelcome.Show();
            this.WindowState = FormWindowState.Maximized;
            //Thread.Sleep(1000);
            //Application.DoEvents();
            InitializeComponent();
            //this.skinEngine1.SkinFile = "Calmness.ssk";
            //sw.Start();
            //listViewToolsLib.RootPath = Application.StartupPath + @"\\Libs";
            //listViewToolsLib.UpdateTools();
            //开启监听键盘和鼠标操作
            Application.AddMessageFilter(new MyIMessageFilter());


            new Thread(() =>
            {
                while (true)
                {
                    try
                    {
                        label85.BeginInvoke(new MethodInvoker(() =>
                            label85.Text = DateTime.Now.ToString()));
                    }
                    catch { }
                    Thread.Sleep(1000);
                }
            })
            { IsBackground = true }.Start();
        }

        string[] sequenceNum = new string[] { };    // 序号
        string[] testItems = new string[] { };      // 测试项目
        string[] actualValue = new string[] { };    // 实际值
        string[] maxValue = new string[] { };       // 上限
        string[] minValue = new string[] { };       // 下限
        string[] beat = new string[] { };           // 节拍   
        string[] unitName = new string[] { };       // 单位
        string[] testResult = new string[] { };     // 测试结果
        string[] standardValue = new string[] { };  // 标准值

        string[] stationName = new string[] { };    // 工位名称
        string[] statisticsCode = new string[] { }; // 生产信息点位集合
        string[] statisticsName = new string[] { }; // 生产信息名称集合

        // 字典用于存储每个CheckBox的初始状态
        private Dictionary<System.Windows.Forms.CheckBox, bool> checkBoxStates = new Dictionary<System.Windows.Forms.CheckBox, bool>();

        Logger loggerAccount = LogManager.GetLogger("AccountManageLog");
        Logger loggerConfig = LogManager.GetLogger("ArgumentConfigLog");

        /// <summary>
        /// 窗体加载
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;

            SearchPort();   // 读卡器设置获取端口

            SYS_Socket_Mo();    // 读取看板界面相关参数

            LoadParameter_MES();// 读取联机参数设置

            SYS_BOARD();        // 读取PLC点位集合

            SYS_Model_Read();   // 读取系统设置

            enterButton_Click(null, null);//加载本地数据源

            this.label54.Text = textBox3.Text;//写设备名称
            FontStyle fontStyle = FontStyle.Bold;//设置字体粗细
            float size = 42F;//字体大小
            changeLabelFont(label54, size, fontStyle);//内字体随着字数的增加而自动减小

            if (textBox30.Text.Length > 0)
            {
                //解析点位字符串成数组
                // actualValue = textBox4.Text.ToString().Split(new char[] { '|' });//采取PLC测试项实际的点
                // testItems = textBox5.Text.ToString().Split(new char[] { '|' });//测试项名称
                // beat = textBox19.Text.ToString().Split(new char[] { '|' });//节拍
                //  maxValue = textBox22.Text.ToString().Split(new char[] { '|' });//采取PLC测试项上限的点
                // minValue = textBox23.Text.ToString().Split(new char[] { '|' });//采取PLC测试项下限的点
                // testResult = textBox29.Text.ToString().Split(new char[] { '|' });//采取PLC测试项结果的点
                // stationCode = textBox31.Text.ToString().Split(new char[] { '|' });//工位点 
                stationName = textBox30.Text.ToString().Split(new char[] { '|' });//工位名称
                                                                                  // sequenceNum = textBox36.Text.ToString().Split(new char[] { '|' });//工位集合
                statisticsCode = textBox9.Text.ToString().Split(new char[] { '|' });
                statisticsName = textBox7.Text.ToString().Split(new char[] { '|' });
            }
            button6_Click(null, null);//用户管理刷新按钮

            //加载打印机
            GetPrint();
            //读取打印信息
            PatchPZLPrinte();

            //修改最后登录时间和次数
            UpLoginInfo();
            try
            {
                ConiferFile coniferFile = ConiferFile.GetJson();
                label52.Text = "版本号: " + coniferFile.Version;
            }
            catch { }

        }

        /// <summary>
        /// 窗体加载后触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Shown(object sender, EventArgs e)
        {
            Process();  // 根据登入信息分配主界面菜单

            // 初始化 CheckBox 的初始状态
            InitializeCheckBoxStates(this.Controls);
            // 配置Nlog全局变量
            LogManager.Configuration.Variables["LoginName"] = LoginName;

            string loginInfo = $"【用户登录】\n工号：{LoginUser} | 姓名：{LoginName} | 权限：{AccessName} | 登录模式：{LoginMode} | 登录方式：{LoginMethod}";
            loggerAccount.Trace(loginInfo);

            try
            {
                string jsonStr = File.ReadAllText(pathText);
                if (!string.IsNullOrWhiteSpace(jsonStr))
                {
                    faultsMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(jsonStr);
                }
            }
            catch { }
            if (faultsMap == null)
            {
                faultsMap = new Dictionary<string, string>();
            }
            if (OffLineType == 0)
            {
                //进行用户登录验证
                Button8_Click(null, null);
            }
            else
            {
                lblRunningStatus.ForeColor = G;
                lblRunningStatus.Text = resources.GetString("UserCheck");
                lblActionTips.Text = resources.GetString("scanning");
            }
            button16_Click(null, null);//连接看板
            TCP_Connect(null, null); //连接PLC
            taskProcess_MES = new Task(Process_MES); //检查PLC状态  写入状态状态信息
            taskProcess_MES.Start();
            Model_Read_Other();  //加载产品型号 

            //根据状态写模式
            //taskProcess_Offline = new Task(Process_Offline);
            //taskProcess_Offline.Start();
            Process_Offline();
            workstNamelist = new List<string>();
            if (sequenceNum.Length > 0)
            {
                workstNamelist = CodeNum.WorkIDName(sequenceNum, stationName);
            }

            richTextBox4.Clear();
            UTYPE.SelectedIndex = 0;
            //判断是否超级管理员、增加定时器
            if (Access == 3)
            {
                timer = new System.Timers.Timer();
                timer.Elapsed += Timer_Elapsed;
                timer.Enabled = true;
                timer.Interval = 1800000;
                timer.Start();
            }

            //初始状态为待机
            lblProductResult.Text = resources.GetString("label_Value");
            lblProductResult.ForeColor = B;
            lblProductResult.BackColor = W;

            //如果联机用户验证失败  直接返回
            if (OffLineType == 0 && loginCheck == false)
            {
                return;
            }
            if (OffLineType == 1)
            {
                textBox6.Text = "111111111111";
            }
            else
            {
                textBox6.Text = Interaction.InputBox(resources.GetString("InputBox"), resources.GetString("InputBoxName"), "", 100, 100);
                for (int i = 1; i <= 5; i++)
                {
                    if (string.IsNullOrWhiteSpace(textBox6.Text))
                    {
                        if (i == 4)
                        {
                            Form1_FormClosed(null, null);
                        }
                        textBox6.Text = Interaction.InputBox(resources.GetString("InputBox") + i + resources.GetString("InputBox1"), resources.GetString("InputBoxName"), "", 100, 100);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            sedbool = true;
            barcodeInfo = "1";
            //Button12_Click(null, null);
            //Button18_Click(null, null);
            //taskMinMaxplc = new Task();//读取上下值数据
            //  taskMinMaxplc.Start();

            taskProcesplc = new Task(Procesplc_ReadCode); //读取条码
            taskProcesplc.Start();
            taskProcesplc1 = new Task(Procesplc_ReadValue);//读取工单号等生产数据
            taskProcesplc1.Start();
            GetMinMaxplc = new Task(Get_MaxMinValue);//读取上下值数据
            GetMinMaxplc.Start();
            taskMinMaxplc = new Task(Bind_MaxMinValue);//绑定上下值数据
            taskMinMaxplc.Start();
            OpenReader_Click(null, null);


            //LineModel_Read();
            if (checkBox7.Checked)
            {
                taskProcess_ZPL = new Task(PrintZPL_Click);//读PLC然后打印
                taskProcess_ZPL.Start();
            }
            if (stationName.Length == 2)
            {
                tabPage10.Text = stationName[0] + stationName[1];
                tabPage11.Text = stationName[0];
                tabPage12.Text = stationName[1];

                InsertTable(null, null);
                InsertTable1(null, null);
                InsertTable2(null, null);
            }
            //strMaxlist = CodeNum.SMaxMindemo(maxValue);
            //strMinlist = CodeNum.SMaxMindemo(minValue);

            //comboBox3.Items.Clear();
            //comboBox3.Items.Add(resources.GetString("Logintype1"));
            //comboBox3.Items.Add(resources.GetString("Logintype2"));
            //comboBox3.SelectedIndex = 0;


        }

        /// <summary>
        /// 递归地初始化所有 CheckBox 的状态，包括嵌套在其他容器中的 CheckBox
        /// </summary>
        /// <param name="controls"></param>
        private void InitializeCheckBoxStates(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                // 如果是CheckBox，记录其状态并绑定CheckedChanged事件
                if (control is System.Windows.Forms.CheckBox cbx)
                {
                    if (!checkBoxStates.ContainsKey(cbx))
                    {
                        checkBoxStates[cbx] = cbx.Checked;
                        checkBox1.CheckedChanged += CheckBox_CheckedChanged;    // 勾选绑定工单
                        checkBox2.CheckedChanged += CheckBox_CheckedChanged;    // 读取PLC
                        checkBox3.CheckedChanged += CheckBox_CheckedChanged;    // 使用文字
                        checkBox4.CheckedChanged += CheckBox_CheckedChanged;    // 读取PLC型号
                        cbxOpenBulletin.CheckedChanged += CheckBox_CheckedChanged;    // 启用看板
                        //checkBox17.CheckedChanged += CheckBox_CheckedChanged;   // PLC控制打印
                        chkBypassBarcodeValidation.CheckedChanged += CheckBox_CheckedChanged;    // 屏蔽本地条码验证
                        checkBox9.CheckedChanged += CheckBox_CheckedChanged;    // 二次读条码
                        chkBypassFixtureVali.CheckedChanged += CheckBox_CheckedChanged;   // 屏蔽本地扫工装验证
                        checkBox11.CheckedChanged += CheckBox_CheckedChanged;   // 屏蔽本地二维码验证
                        //checkBox12.CheckedChanged += CheckBox_CheckedChanged;   // 屏蔽本地NG历史数据
                        //checkBox13.CheckedChanged += CheckBox_CheckedChanged;   // 屏蔽本地历史数据
                    }
                }
                // 如果是容器控件，递归调用此方法
                else if (control is ContainerControl || control is Panel || control is System.Windows.Forms.GroupBox || control is TabPage || control is TabControl)
                {
                    InitializeCheckBoxStates(control.Controls);
                }
            }
        }

        /// <summary>
        /// // CheckedChanged 事件处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (sender is System.Windows.Forms.CheckBox cbx)
            {
                // 检查字典中是否存在对应的键
                if (checkBoxStates.TryGetValue(cbx, out bool initialState))
                {
                    // 当状态发生变化且与初始状态不同
                    if (cbx.Checked != initialState)
                    {
                        // 记录日志
                        string state = cbx.Checked ? "启用" : "已关闭";
                        loggerConfig.Trace($"【{cbx.Text}】{state}");

                        // 更新字典中的状态
                        checkBoxStates[cbx] = cbx.Checked;
                    }
                }
                else
                {
                    loggerConfig.Warn($"未能在字典中找到 CheckBox '{cbx.Name}' 的初始状态。");
                }
            }
        }

        //public void language()
        //{
        //    string language = Properties.Settings.Default.DefaultLanguage;
        //    if (language == "zh-CN")
        //    {
        //        //修改默认语言
        //        MultiLanguage.SetDefaultLanguage("zh-CN");
        //        this.CurrentSelectedLanguage = Language.ChineseSimplified;
        //        //对所有打开的窗口重新加载语言
        //        foreach (Form form in Application.OpenForms)
        //        {
        //            LoadAll(form);
        //        }

        //    }
        //    else if (language == "en-US")
        //    {
        //        //修改默认语言
        //        MultiLanguage.SetDefaultLanguage("en-US");
        //        this.CurrentSelectedLanguage = Language.English;
        //        //对所有打开的窗口重新加载语言
        //        foreach (Form form in Application.OpenForms)
        //        {
        //            LoadAll(form);
        //        }

        //    }
        //    else if (language == "th-TH")
        //    {
        //        //修改默认语言
        //        MultiLanguage.SetDefaultLanguage("th-TH");
        //        this.CurrentSelectedLanguage = Language.Thai;
        //        //对所有打开的窗口重新加载语言
        //        foreach (Form form in Application.OpenForms)
        //        {
        //            LoadAll(form);
        //        }
        //    }
        //}
        //private void LoadAll(Form form)
        //{
        //    MultiLanguage.LoadLanguage(form, typeof(Form1));
        //}

        //达到时间间隔发生的方法
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            iOperCount++;
            if (iOperCount >= 1)
            {
                Console.WriteLine("30分钟未动作程序退出！");
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// 选择打印机
        /// </summary>
        private void GetPrint()
        {
            comboBox1.Items.Add("无");
            foreach (string pkInstalledPrinters in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                cmbInstalledPrinters.Items.Add(pkInstalledPrinters);
                comboBox1.Items.Add(pkInstalledPrinters);
            }
            if (comboBox1.Items.Contains("ZDesigner GK888t (EPL)"))
                if (cmbInstalledPrinters.Items.Contains("ZDesigner GK888t (EPL)"))
                {
                    cmbInstalledPrinters.Text = "ZDesigner GK888t (EPL)";
                    comboBox1.Text = "ZDesigner GK888t (EPL)";
                }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
            //Environment.Exit(0);
        }

        #region --------------- "tools manager" ---------------

        private void StopTools()
        {
            foreach (KeyValuePair<string, IToolBase> tool in mTools)
            {
                tool.Value.IsRunLoop = false;
            }
        }

        private Dictionary<string, IToolBase> mTools = new Dictionary<string, IToolBase>();
        Dictionary<string, IToolBase> IMesDatasBase.Tools
        {
            get
            {
                return mTools;
            }
            set
            {
                mTools = value;
            }
        }
        private SendCommand SendComamandTest;
        public SendCommand sendCommand
        {
            get
            {
                return SendComamandTest;
            }
            set
            {
                SendComamandTest = value;
            }
        }
        // delegate IToolBase delegateLoadTool(string path);
        private IToolBase LoadTool(string toolPath)
        {
            //if (InvokeRequired)
            //{
            //    Invoke(new delegateLoadTool(LoadTool), new object[] { toolPath });
            //    return;
            //}

            // @"D:\YCwork\项目文件\2022\21.深圳比亚迪采集软件\MesDatas\ToolSample\bin\Debug\ToolSample.dll"
            try
            {
                Assembly ass = Assembly.LoadFrom(toolPath);
                var wormMain = ass.GetTypes().FirstOrDefault(m => m.GetInterface(typeof(IToolBase).Name) != null);
                var tmpObj = (IToolBase)Activator.CreateInstance(wormMain);
                if (tmpObj != null)
                {
                    tmpObj.ToolInit();
                    tmpObj.receiveChanged += EnqueueInteracrive;
                    tmpObj.MesDatasMain = this;
                    return tmpObj;
                }
                return null;

            }
            catch
            {
                return null;
            }

        }

        #endregion

        #region --------------- "EnqueueInteracrive" ---------------

        private void EnqueueInteracrive(InteractiveEventArgs e)
        {
            // MessageBox.Show(e.Info);
            Invoke(new Action(() =>
            {
                lock (lockQueue)
                {
                    InterractiveQueue.Enqueue(e);
                }
            }));
        }

        private Queue<InteractiveEventArgs> InterractiveQueue = new Queue<InteractiveEventArgs>();
        Task taskProcesplc = null;
        Task taskMinMaxplc = null;
        Task GetMinMaxplc = null;
        Task taskProcesplc1 = null;
        Task taskPLCNFC = null;
        Task taskProcess = null;
        Task taskProcess_Offline = null;
        Task taskProcess_MES = null;
        Task taskProcess_ZPL = null;
        public bool IsRunningplc = true;
        bool IsRunning = true;
        bool IsRunningplc_ReadCode = true;
        bool IsRunningplc_ReadValue = true;
        bool IsRunningplc_ReadMaxMin = true;
        bool IsRunningplc_MES = true;
        bool IsRunningplc_NFC = true;
        bool IsRunningplc_tabPage = true;
        bool IsRunningPZL = true;
        object lockQueue1 = new object();
        object lockQueue = new object();//pathText
        public static string path4 = System.AppDomain.CurrentDomain.BaseDirectory + "SystemDateBase.mdb";
        public static string pathText = System.AppDomain.CurrentDomain.BaseDirectory + "logfault.txt";
        public static string userFileuRL = "D:\\BYD_Users\\Users_Data.MDB";

        /// 处理交互信息
        /// </summary>
        private void ProcesLoop()
        {
            while (IsRunning)
            {
                InteractiveEventArgs e = null;
                if (InterractiveQueue.Count > 0)
                {
                    lock (lockQueue)
                    {
                        Invoke(new Action(() =>
                        {
                            e = InterractiveQueue.Dequeue();
                        }));

                    }
                    if (e.InfoType == InfoType.Command)
                    {
                        Invoke(new Action(() =>
                        {
                            SendComamandTest?.Invoke(e);
                        }));

                    }
                    else if (e.InfoType == InfoType.Content)
                    {
                        Invoke(new Action(() =>
                        {
                            //update data to grid view
                            UpdateDataToDataGridView(e);
                        }));
                        //update data to grid view

                    }
                    else if (e.InfoType == InfoType.LogMsg)
                    {
                        Invoke(new Action(() =>
                        {
                            //update data to grid view
                            // ShowLog(e.Value);
                        }));
                    }

                }
                Thread.Sleep(5);
                Application.DoEvents();
            }
        }

        private void Process_MES()
        {
            while (IsRunningplc_MES)
            {
                this.Invoke(new Action(() =>
                {
                    if (plcConn == true)
                    {
                        lblPlcStatus.ForeColor = G;
                        //  label115.Text = "PLC状态：已连接";
                    }
                    else
                    {
                        lblPlcStatus.ForeColor = R;
                        //   label115.Text = "PLC状态：未连接";
                    }
                    try
                    {
                        if (!string.IsNullOrEmpty(""))
                        {
                            if (IsStart)
                            {
                                KeyenceMcNet.WriteAsync("", 1);
                            }
                            else
                            {
                                KeyenceMcNet.WriteAsync("", 0);
                            }
                        }
                    }
                    catch (Exception) { }
                    // IsRunningplc_MES = false;
                }));

                Thread.Sleep(100);
                Application.DoEvents();
            }
        }

        private void Process_Offline()
        {
            //while (IsRunning)
            //{
            //    this.Invoke(new Action(() =>
            //    {
            if (OffLineType == 1)
            {

                label19.Text = resources.GetString("loginMode1");
                label41.Text = $"{LoginUser.ToString()} ({LoginName})";
            }
            else if (OffLineType == 0)
            {
                label19.Text = resources.GetString("loginMode");
                label41.Text = $"{LoginUser.ToString()} ({LoginName})";
            }
            //}));

            //Thread.Sleep(100);
            //Application.DoEvents();
            // }
        }

        /// <summary>
        /// 读取条码
        /// </summary>
        private void Procesplc_ReadCode()
        {
            while (IsRunningplc_ReadCode)
            {
                // 触发通讯读
                Button19_Click(null, null);
            }
            Thread.Sleep(50);
            Application.DoEvents();
        }

        /// <summary>
        /// 生产数据读取、上传
        /// </summary>
        private void Procesplc_ReadValue()
        {
            while (IsRunningplc_ReadValue)
            {
                // 触发通讯读
                var short_D1 = KeyenceMcNet.ReadInt16("D1016").Content;
                var short_D2 = KeyenceMcNet.ReadInt16("D1018").Content;

                if (short_D1 == 1 && short_D2 == 0)
                {
                    Invoke(new Action(() =>
                    {
                        tabControl2.SelectedTab = tabPage11;
                        if (checkBox2.Checked == true)
                        {
                            checkBox9.Checked = true;
                        }
                        else
                        {
                            checkBox9.Checked = false;
                        }
                    }));
                    Button20_Click(null, null, "1");
                }
                else if (short_D1 == 0 && short_D2 == 1)
                {
                    Invoke(new Action(() =>
                    {
                        tabControl2.SelectedTab = tabPage12;
                        if (checkBox2.Checked == true)
                        {
                            checkBox9.Checked = true;
                        }
                        else
                        {
                            checkBox9.Checked = false;
                        }
                    }));
                    Button20_Click(null, null, "2");

                }
                else if (short_D1 == 1 && short_D2 == 1)
                {
                    Invoke(new Action(() =>
                    {
                        tabControl2.SelectedTab = tabPage10;
                        checkBox9.Checked = true;
                    }));
                    Button20_Click(null, null, "3");
                }
                Thread.Sleep(100);
                Application.DoEvents();
            }
            Thread.Sleep(50);
            Application.DoEvents();
        }

        private void Process()
        {


            if (Access_take == 1)
            {
                if (Access == 1)//操作员
                {
                    this.tabPage2.Parent = this.tabControl1;
                    this.tabPage3.Parent = this.tabControl1;
                    this.tabPage4.Parent = null;
                    this.tabPage5.Parent = null;
                    this.tabPage6.Parent = null;
                    this.tabPage7.Parent = null;
                    this.tabPage8.Parent = this.tabControl1;
                    this.tabPage9.Parent = null;
                }

                else if (Access == 2)//工艺工程师
                {
                    this.tabPage2.Parent = this.tabControl1;
                    this.tabPage3.Parent = this.tabControl1;
                    this.tabPage4.Parent = this.tabControl1;
                    this.tabPage5.Parent = this.tabControl1;
                    this.tabPage6.Parent = this.tabControl1;
                    this.tabPage7.Parent = null;
                    this.tabPage8.Parent = this.tabControl1;
                    this.tabPage9.Parent = this.tabControl1;
                }

                else if (Access == 3)//管理员（超级用户）
                {
                    this.tabPage2.Parent = this.tabControl1;    // 生产日志
                    this.tabPage3.Parent = this.tabControl1;    // 用户管理
                    this.tabPage4.Parent = this.tabControl1;    // 打印设置
                    this.tabPage5.Parent = this.tabControl1;    // MES参数
                    this.tabPage6.Parent = this.tabControl1;    // 看板设置
                    this.tabPage7.Parent = this.tabControl1;    // 系统设置
                    this.tabPage8.Parent = this.tabControl1;    // 历史数据
                    this.tabPage9.Parent = this.tabControl1;    // 配方设置
                }

                else if (Access == 4)//开发者
                {
                    this.tabPage2.Parent = this.tabControl1;
                    this.tabPage3.Parent = this.tabControl1;
                    this.tabPage4.Parent = this.tabControl1;
                    this.tabPage5.Parent = this.tabControl1;
                    this.tabPage6.Parent = this.tabControl1;
                    this.tabPage7.Parent = this.tabControl1;
                    this.tabPage8.Parent = this.tabControl1;
                    this.tabPage9.Parent = this.tabControl1;
                }

                else if (Access == 5)   // 品质
                {
                    this.tabPage2.Parent = this.tabControl1;
                    this.tabPage3.Parent = this.tabControl1;
                    this.tabPage4.Parent = null;
                    this.tabPage5.Parent = null;
                    this.tabPage6.Parent = null;
                    this.tabPage7.Parent = null;
                    this.tabPage8.Parent = this.tabControl1;
                    this.tabPage9.Parent = null;
                }

                else if (Access == 6)   // 设备
                {
                    this.tabPage2.Parent = this.tabControl1;
                    this.tabPage3.Parent = this.tabControl1;
                    this.tabPage4.Parent = this.tabControl1;
                    this.tabPage5.Parent = null;
                    this.tabPage6.Parent = null;
                    this.tabPage7.Parent = this.tabControl1;
                    this.tabPage8.Parent = this.tabControl1;
                    this.tabPage9.Parent = null;
                }

                Access_take = 0;


            }

        }

        #region---------Label内字体随着字数的增加而自动减小，Label大小不变---------

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
                if (ColNum <= NeedColNum)
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

        NLog.Logger logger1Production = NLog.LogManager.GetLogger("ProductionLog");

        private void LogMsg(string msg)
        {
            this.Invoke(new Action(() =>
            {
                if (richTextBox4.TextLength > 50000)
                {
                    richTextBox4.Clear();
                }

                richTextBox4.AppendText(DateTime.Now.Hour.ToString() +
                    DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" +
                    DateTime.Now.Millisecond.ToString() + ":" + msg + "\r\n");
                richTextBox4.ScrollToCaret();
                // SaveCSVlog(msg);
                logger1Production.Trace(msg);
            }));
        }

        public void SaveCSVlog(string log)
        {
            try
            {
                //  myJobManager.Run();
                if (System.IO.Directory.Exists("D:\\Log") == false)
                {
                    System.IO.Directory.CreateDirectory("D:\\Log");
                }
                // StringBuilder i = new StringBuilder();
                StringBuilder DataLine = new StringBuilder();

                string strT = DateTime.Now.Hour.ToString() + "时" + DateTime.Now.Minute.ToString() + "分" + DateTime.Now.Second.ToString() + "秒";

                //列标题
                // i.Append(log);
                //行数据
                DataLine.Append(strT + ":" + log);
                string FileName = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                string FilePath = "D:\\Log" + "\\" + FileName + ".CSV";

                if (System.IO.File.Exists(FilePath) == false)
                {
                    System.IO.StreamWriter stream = new System.IO.StreamWriter(FilePath, true, Encoding.UTF8);
                    //stream.WriteLine(i);
                    stream.WriteLine(DataLine);
                    stream.Flush();
                    stream.Close();
                    stream.Dispose();
                }
                else
                {
                    System.IO.StreamWriter stream = new System.IO.StreamWriter(FilePath, true, Encoding.UTF8);
                    stream.WriteLine(DataLine);
                    stream.Flush();
                    stream.Close();
                    stream.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return;

        }

        public static bool CreateAccessDatabase(string path)
        {
            //如果文件存在反回假
            if (File.Exists(path))
            {
                MessageBox.Show("文件已存在！");
                return false;
            }

            try
            {
                //如果目录不存在，则创建目录
                string dirName = Path.GetDirectoryName(path);
                if (!Directory.Exists(dirName))
                {
                    Directory.CreateDirectory(dirName);
                }

                //创建Catalog目录类
                ADOX.CatalogClass catalog = new ADOX.CatalogClass();

                string _connectionStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path
                                         + ";";
                //根据联结字符串使用Jet数据库引擎创建数据库
                catalog.Create(_connectionStr);

                //要加上下面这两句，否则创建文件后会有*.ldb文件，一直到程序关闭后
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog.ActiveConnection);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(catalog);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("数据库创建失败:{0}", ex.Message));
            }
        }

        private void UpdateDataToDataGridView(InteractiveEventArgs e)
        {

            //   dataGridView1.UpdateTestData(e.Name, e.Value, e.IsKey);
        }

        #endregion

        private void Form1_Resize(object sender, EventArgs e)
        {
            //  SetLayout();
        }

        #region --------------- "Parameters & Tools Manager" ---------------

        private void SaveTools()
        {

            string tools = null;
            foreach (KeyValuePair<string, IToolBase> tool in mTools)
            {
                tools += tool.Key + "_";
            }
            if (tools != null)
            {
                tools.TrimEnd('_');
                WriteString("Tools", "ToolList", tools);

                foreach (KeyValuePair<string, IToolBase> tool in mTools)
                {
                    tool.Value.SaveParameters();
                }
            }

        }

        private void LoadParam()
        {
            try
            {
                //加载工具
                //加载Config

                //CKL.Checked = Convert.ToBoolean(ReadString("工作模式", "是否联机"));
                //WLCombox.SelectedIndex = Convert.ToInt16(ReadString("物料索引号", "物料"));

                //Laser_add = Convert.ToInt32(ReadString("镭射批号数据", "批号"));
                //LaserString = INIClass.INIGetStringValue(filepath, "镭射字符判定", "记忆字符");


                //plcAddr.Text = ReadString("地址", "PLCIP");
                //plcPort.Text = ReadString("地址", "PLC端口");

                //ccdAddr.Text = ReadString("地址", "ccdIP");
                //ccdPort.Text = ReadString("地址", "ccd端口");

                //prsAddr.Text = ReadString("地址", "扭力IP");
                //prsPort.Text = ReadString("地址", "扭力端口");
                //MesUpdateCmd.Text = ReadString("Commands", "MesUpdateCmd");
                //UserCheckCmd.Text = ReadString("Commands", "UserCheckCmd");
                //OfflineSaveCmd.Text = ReadString("Commands", "OfflineSaveCmd");

                //GridCmds.LoadCacheDataFromCsv();




            }
            catch (Exception)
            {
                // ShowLog(e.ToString());
            }

        }

        private void SaveParam()
        {
            //WriteString("物料索引号", "物料", WLCombox.SelectedIndex.ToString());
            //WriteString("工作模式", "是否联机", CKL.Checked.ToString());

            //WriteString("地址", "PLCIP", plcAddr.Text);
            //WriteString("地址", "PLC端口", plcPort.Text);

            //WriteString("Commands", "MesUpdateCmd", MesUpdateCmd.Text);
            //WriteString("Commands", "UserCheckCmd", UserCheckCmd.Text);
            //WriteString("Commands", "OfflineSaveCmd", OfflineSaveCmd.Text);


            SaveTools();
            //GridCmds.SaveCacheDateToCSV();

        }

        #endregion

        private string filepath = Application.StartupPath + "\\Config.ini";

        private string ReadString(string section, string key)
        {
            return INIClass.INIGetStringValue(filepath, section, key);
        }

        //object lockFileWrite = new object();
        private bool WriteString(string section, string key, string value)
        {
            bool b;
            // lock (lockFileWrite)
            {
                b = INIClass.INIWriteValue(filepath, section, key, value);
            }


            return b;
        }
        object objLockLog = new object();

        #region --------------- 数据查询 ---------------

        private void ButtonSearch_Click(object sender, EventArgs e)
        {
            // DateToAccess();
            // dataGridViewDynamic1.ResetGrid(true);
            //datatable();
            chaxun();
            if (pager != null)
            {
                pager.fenye();//分页
                PageLoad();//显示分页数据
                           //dataset();
            }
        }

        OleDbDataAdapter adp;

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
            //textBoxPath.Text = ReadString("Path");
            //textBoxCmd.Text = ReadString("Command");
        }

        public void SaveParameter()
        {
            // WriteString
            //WriteString("Path", textBoxPath.Text);
            //WriteString("Command", textBoxCmd.Text);
        }

        /// <summary>
        /// 查询
        /// </summary>
        public void chaxun()
        {
            DateTime times_Start = dateTimePicker1.Value;
            DateTime times_End = dateTimePicker2.Value;
            //times_Start= times_Start.ToString("yyyy年MM月dd日 HH:mm:ss");
            // times_End= times_Start.ToString("yyyy年MM月dd日 HH:mm:ss");
            string times_Start_string = times_Start.ToString("yyyy年MM月dd日 HH:mm:ss");
            string times_End_string = times_End.ToString("yyyy年MM月dd日 HH:mm:ss");//dateTimePicker2.Value.ToString("yyyy/MM/dd").ToString();
            DateTime times_Start_Sub = DateTime.Parse(times_Start_string);
            DateTime times_End_Sub = DateTime.Parse(times_End_string);
            if (textBoxPath.Text.Length < 1)
            {
                MessageBox.Show("请先选择左侧数据源！");
                return;
            }
            string dataBaseUrl = textBoxPath.Text + ".mdb";
            mdb = new mdbDatas(dataBaseUrl);

            StringBuilder sql = new StringBuilder();

            sql.Append("select * from [Sheet1] where CDate(Format(测试时间,'yyyy/MM/dd HH:mm:ss')) between  #" + times_Start_Sub + "#" + " and #" + Convert.ToDateTime(times_End_Sub) + "# ");

            if (textBox_Code.Text != "")
            {
                sql.Append("and 条码 like '" + textBox_Code.Text + "%'");
            }
            if (textBox1.Text != "")
            {
                sql.Append("and 产品 = '" + textBox1.Text + "' ");
            }

            DataTable table = mdb.Find(sql.ToString());
            //tablelist_date(dataBaseUrl);
            //dataGridViewDynamic1.SetDataTable(table);
            // dataGridViewDynamic1.DataSource = table;
            pager = new Pager(table);//分页
            mdb.CloseConnection();
        }

        public void tablelist_date(string url)
        {
            mdb = new mdbDatas(url);
            var sql1 = "select * from [Sheet1]";
            DataTable dt = mdb.Find(sql1);
            if (dt.Columns.Count > 0)
            {

                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    dataGridViewDynamic1.AddHeader(dt.Columns[i].ColumnName + "\t");
                }
            }
            mdb.CloseConnection();

        }

        /// <summary>
        /// 返回datatable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void datatable()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";

            OleDbParameter[] pars = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text),
                new OleDbParameter("@column2",dateTimePicker1.ToString())
            };

            var sql = "select * from HCBI开关组 where 条码 = @column1 and 测试时间 = @column2";
            // 条码,测试人,测试时间,测试结果
            DataTable table = AccessHelper.ExecuteDataTable(conn, sql, pars);

            dataGridViewDynamic1.SetDataTable(table);

            //foreach (DataRow row in table.Rows)
            //{
            //    foreach (DataColumn column in table.Columns)
            //    {
            //        listBox1.Text+= (row[column] + "\t");
            //    }
            //}
        }

        /// <summary>
        /// 返回dataSet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void dataset()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            OleDbParameter[] pars = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text),
                //new OleDbParameter("@column2",tbx_Vaule.Text)
            };
            var sql = "select 条码,测试人,测试时间 from HCBI开关组 where 条码 = @column1 "; //and 测试节拍 = @column2
            DataTable table = AccessHelper.ExecuteDataTable(conn, sql, pars);
            OleDbParameter[] pars1 = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text)
            };
            var sql1 = "select 条码,测试人,测试时间 from HCBI开关组 where 条码 = @column1 ";
            DataSet ds = AccessHelper.ExecuteDataSet(conn, sql1, pars1);

            foreach (DataTable tb in ds.Tables)
            {
                foreach (DataColumn col in tb.Columns)
                {
                    dataGridViewDynamic1.AddHeader(col.ColumnName + "\t");
                }

                foreach (DataRow row in table.Rows)
                {
                    foreach (DataColumn col in table.Columns)
                    {
                        //(row[col] + "\t");
                    }

                }
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            SaveParameter();
        }

        #endregion

        #region --------------- MES上传 ---------------
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

        /// <summary>
        /// 加载MES联机参数到系统
        /// </summary>
        public void LoadParameter_MES()
        {
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemInfo where ID = '1'");
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                for (int j = 0; j < table1.Columns.Count; j++)
                {
                    textBox_ip.Text = table1.Rows[i]["IP"].ToString();
                    textBox_port.Text = table1.Rows[i]["Port"].ToString();
                    textBox_timeout.Text = table1.Rows[i]["Timeout"].ToString();
                    textBox_nccode.Text = table1.Rows[i]["NcCode"].ToString();
                    textBox_opration.Text = table1.Rows[i]["Opration"].ToString();
                    textBox_password.Text = table1.Rows[i]["Password"].ToString();
                    textBox_resource.Text = table1.Rows[i]["Resource"].ToString();
                    textBox_site.Text = table1.Rows[i]["Site"].ToString();
                    textBox_url.Text = table1.Rows[i]["Url"].ToString();
                    textBox_user.Text = table1.Rows[i]["User"].ToString();
                    //label41.Text = userCollection.Rows[i]["User"].ToString();
                    //mes cmd
                    //UserCheckCmd.Text = userCollection.Rows[i]["UserCheckCmd"].ToString();
                    //UserCheckPass.Text = userCollection.Rows[i]["UserCheckPass"].ToString();
                    //UserCheckFail.Text = userCollection.Rows[i]["UserCheckFail"].ToString();
                    //CodeCheckCmd.Text = userCollection.Rows[i]["CodeCheckCmd"].ToString();
                    //CodeCheckPass.Text = userCollection.Rows[i]["CodeCheckPass"].ToString();
                    //CodeSendCmd.Text = userCollection.Rows[i]["CodeSendCmd"].ToString();
                    //FileVersion.Text = userCollection.Rows[i]["FileVersion"].ToString();
                    //SoftwareVersion.Text = userCollection.Rows[i]["SoftwareVersion"].ToString();
                }
            }
            mdb.CloseConnection();
        }

        /// <summary>
        /// 保存MES联机参数到数据库
        /// </summary>
        public void SaveParameter_MES()
        {
            SytemInfoEntity infoEntity = new SytemInfoEntity();
            infoEntity.IP = textBox_ip.Text;
            infoEntity.Port = textBox_port.Text;
            infoEntity.Timeout = textBox_timeout.Text;
            infoEntity.NcCode = textBox_nccode.Text;
            infoEntity.Opration = textBox_opration.Text;
            infoEntity.Password = textBox_password.Text;
            infoEntity.Resource = textBox_resource.Text;
            infoEntity.Site = textBox_site.Text;
            infoEntity.Url = textBox_url.Text;
            infoEntity.User = textBox_user.Text;
            //infoEntity.FileVersion = FileVersion.Text;
            //infoEntity.SoftwareVersion = SoftwareVersion.Text;
            //infoEntity.UserCheckCmd = "";
            //infoEntity.UserCheckPass = UserCheckPass.Text;
            //infoEntity.UserCheckFail = UserCheckFail.Text;
            //infoEntity.CodeCheckCmd = CodeCheckCmd.Text;
            //infoEntity.CodeCheckPass = CodeCheckPass.Text;
            //infoEntity.CodeSendCmd = CodeSendCmd.Text;

            saveSystemInfo(path4, infoEntity);
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

        public void 绑定工单_Mes(string 工单号, out bool 绑定结果, out string MES反馈, out string XMLOUT)
        {
            BydMesCom.绑定工单(工单号, out 绑定结果, out MES反馈, out XMLOUT);
        }

        bool loginCheck = false;

        /// <summary>
        /// 用户登录验证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button8_Click(object sender, EventArgs e)
        {
            lblRunningStatus.Text = resources.GetString("user_yzz");
            lblActionTips.Text = resources.GetString("Wait");
            Config_Mes(ip, port, timeout, url, site, user, password, resource, operation, nccode);
            bool 验证结果;
            string MES反馈;
            string XMLOUT;
            UsersVarify(out 验证结果, out MES反馈, out XMLOUT);
            if (验证结果 == true)
            {
                if (MES反馈 != null)
                {
                    richTextBox1.Clear();
                    richTextBox1.AppendText(MES反馈);
                    lblRunningStatus.ForeColor = G;
                    lblRunningStatus.Text = resources.GetString("onlineUser_OK");
                    lblActionTips.Text = resources.GetString("scanning");
                    loginCheck = true;
                }
                else
                {
                    richTextBox1.Clear();
                    richTextBox1.AppendText(MES反馈);
                    lblRunningStatus.ForeColor = R;
                    lblRunningStatus.Text = resources.GetString("onlineUser_NG");
                    lblActionTips.Text = resources.GetString("Check_param");
                }
            }
            else
            {
                richTextBox1.Clear();
                richTextBox1.AppendText(MES反馈);
                lblRunningStatus.ForeColor = R;
                lblRunningStatus.Text = resources.GetString("onlineUser_NG");
                lblActionTips.Text = resources.GetString("Check_param");
            }
        }

        NLog.Logger loggerMESBarCoode = NLog.LogManager.GetLogger("MESBarCoodeLog");

        /// <summary>
        /// 条码验证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button9_Click(object sender, EventArgs e)
        {
            Parameter_txt[2002] = "0";
            Parameter_txt[2004] = "0";
            string 产品条码 = barcodeInfo;
            bool 验证结果;
            string MES反馈;
            string XMLOUT; ;
            BarCodeVarify(产品条码, out 验证结果, out MES反馈, out XMLOUT);
            richTextBox1.Clear();
            richTextBox1.AppendText(MES反馈);
            loggerMESBarCoode.Trace(MES反馈);
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            richTextBox1.ScrollToCaret();
            if (验证结果 == true)
            { Parameter_txt[2002] = "1"; }
            else
            {
                Parameter_txt[2004] = "1";
            }
        }

        NLog.Logger loggerMESData = NLog.LogManager.GetLogger("MESDataLog");

        /// <summary>
        /// 结果上传
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button7_Click(object sender, EventArgs e)
        {
            Parameter_txt[2006] = "0";
            Parameter_txt[2008] = "0";
            LogMsg("结果联机上传");
            bool 测试结果 = 产品结果; string 产品条码 = barcodeData; string 文件版本 = Parameter_Model[17500]; ; string 软件版本 = Parameter_Model[17800];
            //string 测试项 = "!测试项,测试参数,测试值"+"!用户ID,"+ label41.Text + ","+ label41.Text + "!测试人,测试时间,测试结果" + "!" +
            //  label41.Text + "," + Runtime.ToString() + "," + Value[9999] + /*"!高度上限,高度下限,当前高度" + "!" + Parameter_txt[1030] + "," + Parameter_txt[1032] + "," + Parameter_txt[1034] +*/
            //"!工装,工装代码," + Parameter_Model[19500] + "!产品编码,产品名称,产品代码" + "!" + Parameter_Model[18000] + "," + Parameter_Model[18500] + "," + Parameter_Model[19000] +
            //"!生产节拍,文件版本,软件版本" + "!" + Parameter_Model[1090] + "," + Parameter_Model[17500] + "," + Parameter_Model[17800];

            StringBuilder sbl = new StringBuilder();
            string[] froddMes = textBox41.Text.Split('+');
            if (froddMes.Length > 0)
            {
                for (int i = 0; i < froddMes.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(froddMes[i]))
                    {
                        sbl.Append("!工装编号" + (i + 1) + ",工装," + froddMes[i]);
                    }
                }
            }
            string[] frdeMafrMes = CodeNum.CodeMafror(comboBox2.Text, codesDataM);
            if (frdeMafrMes.Length > 0)
            {
                for (int i = 0; i < frdeMafrMes.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(frdeMafrMes[i]))
                    {
                        sbl.Append("!产品物料号" + (i + 1) + ",物料," + frdeMafrMes[i]);
                    }
                }
            }
            if (actualValue.Length > 0)
            {
                for (int i = 0; i < list.Count; i++)
                {
                    if (!list[i].Equals("null"))
                    {

                        if (!maxlist[i].Equals("null") && !minlist[i].Equals("null"))
                        {
                            sbl.Append("!" + AstrName[i] + "," + maxlist[i] + "~" + minlist[i] + "," + list[i] + "");
                            sbl.Append("!" + AstrName[i] + "测试结果,结果," + resultlist[i]);
                        }
                        else
                        {
                            sbl.Append("!" + AstrName[i] + ",测试项值," + list[i] + "");
                        }

                    }
                }
            }

            string 测试项 = "!用户ID," + LoginUser.ToString() + "," + LoginUser.ToString() + "!条码,条码信息," + 产品条码 +
                "!产品型号,型号," + tbxProductName.Text +
                "" + sbl.ToString() + "!测试总结果,测试结果," + Value[9999];

            //!测试项,测试参数,测试值!用户ID,YC,YC!测试人,测试时间,测试结果!YC,2023年8月2日19：1：3,OK!工装,工装代码,41!产品编码,产品名称,产品代码!1144_00,SA3F_5820120,0SF5!生产节拍,文件版本,软件版本!5,20220811,20220811
            bool 验证结果; string MES反馈; string XMLOUT;
            UpDateToMes(测试结果, 产品条码, 文件版本, 软件版本, 测试项, out 验证结果, out MES反馈, out XMLOUT);
            richTextBox1.Clear();
            richTextBox1.AppendText(MES反馈);
            loggerMESData.Trace(MES反馈);
            //richTextBox2.AppendText(测试项);
            if (验证结果 == true)
            {
                Parameter_txt[2006] = "1";
            }
            else
            {
                Parameter_txt[2008] = "1";
            }
            LogMsg("联机结果上传完成" + "联机上传结果(true/false)" + 验证结果);

        }

        private void 绑定工单(object sender, EventArgs e)
        {
            绑定工单_Mes(工单号, out bool 绑定结果, out string MES反馈, out string XMLOUT);
            //LogMsg(MES反馈);
        }

        //private static ModbusTcpNet KeyenceMcNet;
        private static KeyenceMcNet KeyenceMcNet;
        private bool plcConn = false;

        private void TCP_Connect(object sender, EventArgs e)
        {
            try
            {
                //KeyenceMcNet = new ModbusTcpNet("127.0.0.1", 503);
                KeyenceMcNet = new KeyenceMcNet(tbx_IP.Text, Convert.ToInt16(tbx_port.Text));
                KeyenceMcNet.ConnectClose();
                KeyenceMcNet.SetPersistentConnection();
                OperateResult connect = KeyenceMcNet.ConnectServer();
                if (connect.IsSuccess)
                {
                    plcConn = true;
                    //MessageBox.Show("PLC连接成功");
                }
                else
                {
                    plcConn = false;
                    MessageBox.Show(resources.GetString("plcConn"));
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }


        #region ---------------------------系统参数设置-----------------------------

        /// <summary>
        /// 保存系统设置
        /// </summary>
        private void SYS_Model_Write()
        {
            // PLC连接参数不能为空
            if (tbx_IP.Text == String.Empty || tbx_port.Text == String.Empty || textBox3.Text == String.Empty)
            {
                MessageBox.Show("当前界面内容均为必填项、请先填写完善");
                return;
            }

            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");

            // 修改
            if (table1.Rows.Count > 0)
            {
                DataRow row = table1.Rows[0];

                // 记录修改前的数据
                string logDetail = $"【修改前的详细信息如下】\n\n" +
                    $"PLC-IP地址：{row["IP"]} | PLC-端口号：{row["Port"]}\n\n" +
                    $"生产数据存放路径：{row["DataUrl"]}\n\n" +
                    $"设备名称：{row["DeviceName"]}\n\n" +
                    $"RFID端口号：{row["rfidProt"]} | RFID设备号：{row["rfidCode"]}\n\n" +
                    $"二次读取产品条码（适用于转盘机台)——勾选状态{row["readBarCode"]}\n\n" +
                    $"生产信息名称集合：\n{row["StatisticsName"]}\n\n" +
                    $"生产信息点位集合：\n{row["StatisticsCode"]}";

                string sql = $"update [SytemSet] set [IP] = '{tbx_IP.Text}', [Port] = '{tbx_port.Text}', " +
                    $"[DataUrl] = '{label2.Text}', [DeviceName] = '{textBox3.Text}', [rfidProt] = '{cmbShowPort.Text}', " +
                    $"[rfidCode] = '{tbxReaderDeviceID.Text}', [StatisticsCode] = '{textBox9.Text}', [StatisticsName] = '{textBox7.Text}', " +
                    $"[ResultCode] = '{textBox5.Text}', [readBarCode] = '{checkBox9.Checked}' where [ID] = '1'";

                var result = mdb.Change(sql);
                if (result)
                {
                    MessageBox.Show("保存成功");

                    string modifyInfo = $"【修改后的详细信息如下】\n\n" +
                        $"PLC-IP地址：{tbx_IP.Text} | PLC-端口号：{tbx_port.Text}\n\n" +
                        $"生产数据存放路径：{label2.Text}\n\n" +
                        $"设备名称：{textBox3.Text}\n\n" +
                        $"RFID端口号：{cmbShowPort.Text} | RFID设备号：{tbxReaderDeviceID.Text}\n\n" +
                        $"二次读取产品条码（适用于转盘机台)——勾选状态：{checkBox9.Checked}\n\n" +
                        $"生产信息名称集合：\n{textBox7.Text}\n\n" +
                        $"生产信息点位集合：\n{textBox9.Text}";
                    loggerConfig.Trace($"【系统参数修改成功】\n\n{logDetail}\n{modifyInfo}");
                }
            }

            // 新增
            else
            {
                mdbDatas.CreateAccessDatabase(path4);
                mdbDatas.CreateMDBTable(path4, "SytemSet", new System.Collections.ArrayList(new object[] { "ID", "IP", "Port", "DataUrl", "DeviceName", "stationCode", "stationName", "wordNo" }));
                DataTable dt = new DataTable("SytemSet");
                DataColumn ID = new DataColumn("ID", typeof(string));
                dt.Columns.Add(ID);
                DataColumn IP = new DataColumn("IP", typeof(string));
                dt.Columns.Add(IP);
                DataColumn Port = new DataColumn("Port", typeof(string));
                dt.Columns.Add(Port);
                DataColumn DataUrl = new DataColumn("DataUrl", typeof(string));
                dt.Columns.Add(DataUrl);
                DataColumn sysName = new DataColumn("DeviceName", typeof(string));
                dt.Columns.Add(sysName);
                DataColumn stationCode = new DataColumn("stationCode", typeof(string));
                dt.Columns.Add(stationCode);
                DataColumn stationName = new DataColumn("stationName", typeof(string));
                dt.Columns.Add(stationName);
                DataColumn wordNo = new DataColumn("wordNo", typeof(string));
                dt.Columns.Add(wordNo);

                DataRow dr = dt.NewRow();
                dt.Rows.Add(dr);

                dr[0] = "1";
                dr[1] = tbx_IP.Text;
                dr[2] = tbx_port.Text;
                dr[3] = label2.Text;
                dr[4] = textBox3.Text;
                // dr[5] = textBox4.Text;
                // dr[6] = textBox5.Text;
                dr[7] = textBox6.Text;
                mdb.DatatableToMdb("SytemSet", dt);
            }

            mdb.CloseConnection();
        }

        /// <summary>
        /// 读取系统设置参数
        /// </summary>
        private void SYS_Model_Read()
        {
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");

            for (int i = 0; i < table1.Rows.Count; i++)
            {
                for (int j = 0; j < table1.Columns.Count; j++)
                {
                    tbx_IP.Text = table1.Rows[i]["IP"].ToString();
                    tbx_port.Text = table1.Rows[i]["Port"].ToString();
                    label2.Text = table1.Rows[i]["DataUrl"].ToString();
                    textBox3.Text = table1.Rows[i]["DeviceName"].ToString();
                    // textBox4.Text = userCollection.Rows[i]["stationCode"].ToString();
                    //textBox5.Text = userCollection.Rows[i]["stationName"].ToString();
                    textBox6.Text = table1.Rows[i]["wordNo"].ToString();
                    //  textBox22.Text = userCollection.Rows[i]["BoardMaxCode"].ToString();//大限度
                    //  textBox23.Text = userCollection.Rows[i]["BoardMinCode"].ToString();//小限度
                    string device = table1.Rows[i]["Workstname"].ToString();
                    if (device == "True") checkBox6.Checked = true;
                    textBox41.Text = table1.Rows[i]["BoardBeat"].ToString();
                    textBox5.Text = table1.Rows[i]["ResultCode"].ToString();//faults
                    comboBox2.SelectedValue = table1.Rows[i]["faults"].ToString();
                    cmbShowPort.SelectedItem = table1.Rows[i]["rfidProt"].ToString();
                    tbxReaderDeviceID.Text = table1.Rows[i]["rfidCode"].ToString();
                    checkBox9.Checked = bool.Parse(table1.Rows[i]["readBarCode"].ToString());
                    //textBox39.Text = userCollection.Rows[i]["faults"].ToString();
                    textBox9.Text = table1.Rows[i]["StatisticsCode"].ToString();
                    textBox7.Text = table1.Rows[i]["StatisticsName"].ToString();
                }
            }
            mdb.CloseConnection();
        }

        #endregion

        #region----------------上下限数据-----------

        DataTable maxminTable = null;

        public void Bind_MaxMinValue()
        {
            while (IsRunningplc_ReadMaxMin)
            {
                PLCMaxMin();
                Thread.Sleep(1800);
                Application.DoEvents();
            }
        }

        /// <summary>
        /// 实时读取上限值
        /// </summary>
        public void Get_MaxMinValue()
        {
            //new Thread(() => {
            while (IsRunningplc_ReadMaxMin)
            {
                //Task.Run(() =>
                // {
                GetPLCMaxMin();
                // }).Wait();
                Thread.Sleep(1500);
                Application.DoEvents();
            }
            //})
            //{ IsBackground = true }.Start();
        }

        public List<MaxMinValue> maxMinValues = null;

        private void GetPLCMaxMin()
        {
            if (plcConn == true)
            {
                maxMinValues = new List<MaxMinValue>();
                for (int i = 0; i < boardTable.Rows.Count; i++)
                {
                    Console.WriteLine(i);
                    MaxMinValue value = new MaxMinValue();

                    value.BoardName = boardTable.Rows[i]["BoardName"].ToString();
                    value.StandardCode = NullModify(PLCCodeNum(boardTable.Rows[i]["StandardCode"].ToString()));
                    value.MaxBoardCode = NullModify(PLCCodeNum(boardTable.Rows[i]["MaxBoardCode"].ToString()));
                    value.MinBoardCode = NullModify(PLCCodeNum(boardTable.Rows[i]["MinBoardCode"].ToString()));
                    value.BoardCode = NullModify(PLCCodeNum(boardTable.Rows[i]["BoardCode"].ToString()));
                    value.Result = NullModify(PLCCodeNum(boardTable.Rows[i]["ResultBoardCode"].ToString()));
                    if (value.BoardCode.Equals("OK") || value.BoardCode.Equals("NG"))
                    {
                        value.Result = value.BoardCode;
                    }
                    maxMinValues.Add(value);
                }
                maxMinList = maxMinValues;
            }
        }

        public string NullModify(string str)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                str = "/";
            }
            if (str.Equals("null"))
            {
                str = "/";
            }
            return str;
        }
        private List<MaxMinValue> maxMinList = null;
        private void PLCMaxMin()
        {
            BeginInvoke(new Action(() =>
            {
                if (plcConn == true)
                {
                    try
                    {
                        List<MaxMinValue> maxMinList = this.maxMinList;
                        if (maxminTable == null)
                        {
                            maxminTable = new DataTable();
                            maxminTable.Columns.Add("序号", typeof(string));
                            maxminTable.Columns.Add("测试项目", typeof(string));
                            maxminTable.Columns.Add("标准值", typeof(string));
                            maxminTable.Columns.Add("上限值", typeof(string));
                            maxminTable.Columns.Add("下限值", typeof(string));
                            maxminTable.Columns.Add("实际值", typeof(string));
                            maxminTable.Columns.Add("测试结果", typeof(string));

                            if (boardTable.Rows.Count > 0)
                            {
                                for (int i = 0; i < boardTable.Rows.Count; i++)
                                {
                                    DataRow dr = maxminTable.NewRow();
                                    dr["序号"] = (i + 1);
                                    dr["测试项目"] = boardTable.Rows[i]["BoardName"].ToString();
                                    dr["标准值"] = (boardTable.Rows[i]["StandardCode"].ToString());
                                    dr["上限值"] = (boardTable.Rows[i]["MaxBoardCode"].ToString());
                                    dr["下限值"] = (boardTable.Rows[i]["MinBoardCode"].ToString());
                                    dr["实际值"] = (boardTable.Rows[i]["BoardCode"].ToString());

                                    if (dr["实际值"].Equals("NG") || dr["实际值"].Equals("OK"))
                                    {
                                        dr["测试结果"] = dr["实际值"];
                                    }
                                    else
                                    {
                                        dr["测试结果"] = (boardTable.Rows[i]["ResultBoardCode"].ToString());
                                    }
                                    maxminTable.Rows.Add(dr);
                                }
                            }

                            dataGridViewDynamic6.DataSource = maxminTable;
                            dataGridViewDynamic6.Columns[1].Width = 280;
                            for (int i = 0; i < dataGridViewDynamic6.Columns.Count; i++)
                            {
                                dataGridViewDynamic6.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                            }
                        }
                        else
                        {
                            if (boardTable.Rows.Count > 0 && dataGridViewDynamic6.Rows.Count > 0)
                            {
                                for (int i = 0; i < dataGridViewDynamic6.Rows.Count; i++)
                                {
                                    if (maxMinList != null && maxMinList.Count > 0 && maxMinList.Count == dataGridViewDynamic6.Rows.Count)
                                    {
                                        dataGridViewDynamic6.Rows[i].Cells[0].Value = (i + 1).ToString();
                                        dataGridViewDynamic6.Rows[i].Cells[1].Value = boardTable.Rows[i]["BoardName"].ToString();
                                        dataGridViewDynamic6.Rows[i].Cells[2].Value = maxMinList[i].StandardCode.ToString();//PLCCodeNum(boardTable.Rows[i]["StandardCode"].ToString());
                                        dataGridViewDynamic6.Rows[i].Cells[3].Value = maxMinList[i].MaxBoardCode.ToString();//PLCCodeNum(boardTable.Rows[i]["MaxBoardCode"].ToString());
                                        dataGridViewDynamic6.Rows[i].Cells[4].Value = maxMinList[i].MinBoardCode.ToString();//PLCCodeNum(boardTable.Rows[i]["MinBoardCode"].ToString());
                                        dataGridViewDynamic6.Rows[i].Cells[5].Value = maxMinList[i].BoardCode.ToString();// PLCCodeNum(boardTable.Rows[i]["BoardCode"].ToString());
                                        dataGridViewDynamic6.Rows[i].Cells[6].Value = maxMinList[i].Result.ToString();
                                        //string shujuViewValew5 = dataGridViewDynamic6.Rows[i].Cells[5].Value.ToString();
                                        //if (shujuViewValew5.Equals("NG") || shujuViewValew5.Equals("OK"))
                                        //{
                                        //    dataGridViewDynamic6.Rows[i].Cells[6].Value = shujuViewValew5;
                                        //}
                                        //else
                                        //{
                                        //    dataGridViewDynamic6.Rows[i].Cells[6].Value = maxMinList[j].Result.ToString();// PLCCodeNum(boardTable.Rows[i]["ResultBoardCode"].ToString());
                                        //}
                                        string str = dataGridViewDynamic6.Rows[i].Cells["测试结果"].Value.ToString();
                                        if (str.Equals("OK"))
                                        {
                                            dataGridViewDynamic6.Rows[i].Cells["测试结果"].Style.BackColor = Color.Green;
                                        }
                                        else if (str.Equals("NG"))
                                        {
                                            dataGridViewDynamic6.Rows[i].Cells["测试结果"].Style.BackColor = Color.Red;
                                        }
                                        else
                                        {
                                            dataGridViewDynamic6.Rows[i].Cells["测试结果"].Style.BackColor = Color.White;
                                        }

                                    }
                                    //dataGridViewDynamic4.Rows[i].Cells[0].Value = (i + 1);
                                    //dataGridViewDynamic4.Rows[i].Cells[1].Value = boardTable.Rows[i]["BoardName"].ToString();
                                    //dataGridViewDynamic4.Rows[i].Cells[2].Value = PLCCodeNum(boardTable.Rows[i]["StandardCode"].ToString());
                                    //dataGridViewDynamic4.Rows[i].Cells[3].Value = PLCCodeNum(boardTable.Rows[i]["MaxBoardCode"].ToString());
                                    //dataGridViewDynamic4.Rows[i].Cells[4].Value = PLCCodeNum(boardTable.Rows[i]["MinBoardCode"].ToString());
                                    //dataGridViewDynamic4.Rows[i].Cells[5].Value = PLCCodeNum(boardTable.Rows[i]["BoardCode"].ToString());
                                    //string shujuViewValew5 = dataGridViewDynamic4.Rows[i].Cells[5].Value.ToString();
                                    //if (shujuViewValew5.Equals("NG") || shujuViewValew5.Equals("OK"))
                                    //{
                                    //    dataGridViewDynamic4.Rows[i].Cells[6].Value = shujuViewValew5;
                                    //}
                                    //else
                                    //{
                                    //    dataGridViewDynamic4.Rows[i].Cells[6].Value = PLCCodeNum(boardTable.Rows[i]["ResultBoardCode"].ToString());
                                    //}
                                    //string str = dataGridViewDynamic4.Rows[i].Cells["测试结果"].Value.ToString();
                                    //if (str.Equals("OK"))
                                    //{
                                    //    dataGridViewDynamic4.Rows[i].Cells["测试结果"].Style.BackColor = Color.Green;
                                    //}
                                    //else if (str.Equals("NG"))
                                    //{
                                    //    dataGridViewDynamic4.Rows[i].Cells["测试结果"].Style.BackColor = Color.Red;
                                    //}
                                    ////maxminTable.Rows[i]["序号"] = (i + 1);
                                    ////maxminTable.Rows[i]["测试项目"] = boardTable.Rows[i]["BoardName"].ToString();
                                    ////maxminTable.Rows[i]["标准值"] = PLCCodeNum(boardTable.Rows[i]["StandardCode"].ToString());
                                    ////maxminTable.Rows[i]["上限值"] = PLCCodeNum(boardTable.Rows[i]["MaxBoardCode"].ToString());
                                    ////maxminTable.Rows[i]["下限值"] = PLCCodeNum(boardTable.Rows[i]["MinBoardCode"].ToString());
                                    ////maxminTable.Rows[i]["实际值"] = PLCCodeNum(boardTable.Rows[i]["BoardCode"].ToString());
                                    ////if (maxminTable.Rows[i]["实际值"].Equals("NG") || maxminTable.Rows[i]["实际值"].Equals("OK"))
                                    ////{
                                    ////    maxminTable.Rows[i]["测试结果"] = maxminTable.Rows[i]["实际值"];
                                    ////}
                                    ////else
                                    ////{
                                    ////    maxminTable.Rows[i]["测试结果"] = PLCCodeNum(boardTable.Rows[i]["ResultBoardCode"].ToString());
                                    ////}
                                }

                            }
                        }

                        //dataGridViewDynamic4.DataSource = maxminTable;

                    }
                    catch { }
                }
            })).AsyncWaitHandle.WaitOne();
        }

        #endregion

        private string proCode = " ";
        private string D1080 = " ";
        private string D1084 = " ";
        private string D1086 = " ";
        private string D1088 = " ";
        private string D1090 = " ";
        private string D1082 = " ";
        private string D1076 = " ";

        string complte = "  ";          // 完成率
        string Paate = "  ";            // 合格率
        string prodproces = "   ";      // 工序时间
        string ustim = "  ";            // 利用时间
        string loadti = "  ";           // 负荷时间
        string shootthgh = "  ";        // 直通率
        string productName = " ";       // 成品名称
        private string runstate = null; // 设备状态

        Dictionary<string, string> faultsMap = new Dictionary<string, string>();
        FaultsBLL faultsBLL = new FaultsBLL();
        short[] faulttime = new short[] { };
        bool frockbool = true;
        DataTable statisticsTable = null;

        /// <summary>
        /// 读取产品型号
        /// </summary>
        private void Model_Read_Other()
        {

            Task.Run(() =>
            {
                while (true)
                {
                    if (plcConn == true)
                    {
                        try
                        {
                            //    this.Invoke((EventHandler)delegate
                            //    {
                            runstate = KeyenceMcNet.ReadInt32("D1007").Content.ToString();//读取设备状态
                            string pCode = KeyenceMcNet.ReadString("D1120", 10).Content;//型号
                            List<string> statislist = StatIsRead();

                            D1080 = statislist[0];// KeyenceMcNet.ReadInt32("D1080").Content.ToString();//生产总数
                            D1084 = statislist[1];// KeyenceMcNet.ReadInt32("D1084").Content.ToString();//工单数量
                            D1086 = statislist[2];// KeyenceMcNet.ReadInt32("D1086").Content.ToString();//完成数量
                            D1088 = statislist[3];// KeyenceMcNet.ReadInt32("D1088").Content.ToString();//合格数量
                            D1090 = statislist[4];//KeyenceMcNet.ReadInt32("D1090").Content.ToString();//生产节拍
                            D1082 = statislist[5];// KeyenceMcNet.ReadInt32("D1082").Content.ToString();//保养计数
                            D1076 = statislist[6];//KeyenceMcNet.ReadInt32("D1076").Content.ToString();//NG数量
                                                  //
                            shootthgh = statislist[7];// CodeNum.PNumCode(KeyenceMcNet.ReadInt32("D1064").Content.ToString());//直通率
                            complte = statislist[8];// CodeNum.PNumCode(KeyenceMcNet.ReadInt32("D1066").Content.ToString());//完成率
                            Paate = statislist[9];// CodeNum.PNumCode(KeyenceMcNet.ReadInt32("D1068").Content.ToString());//合格率 
                            prodproces = statislist[10];// CodeNum.PNtimeCode(KeyenceMcNet.ReadInt32("D1070").Content.ToString());//工序时间
                            ustim = statislist[11];// CodeNum.PNtimeCode(KeyenceMcNet.ReadInt32("D1072").Content.ToString());//利用时间
                            loadti = statislist[12];// CodeNum.PNtimeCode(KeyenceMcNet.ReadInt32("D1074").Content.ToString());//负荷时间
                            proCode = CodeNum.StrVBcd(pCode);
                            faulttime = KeyenceMcNet.ReadInt16(textBox39.Text, CodeNum.Shoubtowule(textBox40.Text)).Content;
                            Invoke(new Action(() =>
                            {

                                if (checkBox6.Checked)
                                {//配方号
                                    string codid = KeyenceMcNet.ReadInt32("D1208").Content.ToString();
                                    comboBox2.SelectedValue = codid;
                                    label50.Text = codid;
                                    string[] frockA = CodeNum.Confrock(comboBox2.Text);
                                    string[] frockB = textBox41.Text.ToString().Split('+');
                                    if (!CodeNum.CopareArr(frockA, frockB))
                                    {
                                        string frockmm = "";
                                        foreach (var frockid in frockB)
                                        {
                                            if (frockA.Contains(frockid))
                                            {
                                                if (string.IsNullOrWhiteSpace(frockmm))
                                                {
                                                    frockmm = frockid;
                                                }
                                                else
                                                {
                                                    frockmm = frockmm + "+" + frockid;
                                                }
                                            }
                                        }
                                        textBox41.Text = frockmm;
                                        string[] frockC = textBox41.Text.ToString().Split('+');
                                        if (!CodeNum.CopareArr(frockA, frockC))
                                        {
                                            lblActionTips.Text = resources.GetString("Wait_scan_Jig");
                                        }
                                    }

                                }

                                textBox42.Text = CodeNum.CodeStrfror(comboBox2.Text, codesDataM);
                                if (runstate != null)
                                {
                                    switch (runstate)//对设备状态进行判断
                                    {
                                        case "1"://1=生产运行 
                                            lblDeviceStatus.ForeColor = G;
                                            // label38.Text = "设备状态：运行";
                                            Fouddinog();
                                            break;
                                        case "2"://2=故障未停机  
                                            lblDeviceStatus.ForeColor = R;
                                            //  label38.Text = "设备状态：触发故障";
                                            Namestation("触发故障");
                                            break;
                                        case "3":// 3=故障停机 
                                            lblDeviceStatus.ForeColor = R;
                                            //  label38.Text = "设备状态：故障停机";
                                            //faultsID.Split[]
                                            Namestation("故障停机");
                                            break;
                                        case "4"://4=待机
                                            lblDeviceStatus.ForeColor = O;
                                            // label38.Text = "设备状态：待机";
                                            Fouddinog();
                                            break;
                                        default:
                                            lblDeviceStatus.ForeColor = R;
                                            // label38.Text = "设备状态：断线";
                                            break;
                                    }
                                }
                                tbxProductName.Text = proCode;
                                if (statisticsName.Length == statisticsCode.Length && statisticsCode.Length > 0)
                                {
                                    if (statisticsTable == null)
                                    {
                                        statisticsTable = new DataTable();
                                        statisticsTable.Columns.Add("名称", typeof(string));
                                        statisticsTable.Columns.Add("值", typeof(string));
                                        for (int i = 0; i < statisticsName.Length; i++)
                                        {
                                            DataRow dr = statisticsTable.NewRow();
                                            dr["名称"] = statisticsName[i];
                                            dr["值"] = statislist[i];
                                            statisticsTable.Rows.Add(dr);
                                        }
                                        dataGridViewDynamic5.DataSource = statisticsTable;
                                        for (int i = 0; i < dataGridViewDynamic5.Columns.Count; i++)
                                        {
                                            dataGridViewDynamic5.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                                        }
                                    }
                                    else
                                    {
                                        for (int i = 0; i < statisticsName.Length; i++)
                                        {
                                            //statisticsTable.Rows[i]["名称"] = statisticsName[i];
                                            //statisticsTable.Rows[i]["值"] = statislist[i];
                                            dataGridViewDynamic5.Rows[i].Cells[0].Value = statisticsName[i];
                                            dataGridViewDynamic5.Rows[i].Cells[1].Value = statislist[i];
                                        }
                                    }

                                }

                                // 从PLC读取条码的产品的名称
                                if (checkBox4.Checked)
                                {
                                    textBox28.Text = tbxProductName.Text;
                                }

                                Thread.Sleep(150);
                                Application.DoEvents();
                            }));
                        }
                        catch { }
                    }
                }
            });
        }

        private List<string> StatIsRead()
        {
            List<string> statislist = new List<string>();
            if (statisticsCode.Length > 0)
            {

                for (int i = 0; i < statisticsCode.Length; i++)
                {
                    if (statisticsCode[i].Contains("-"))
                    {
                        statislist.Add(PLCCodeNum(statisticsCode[i]));
                    }
                    else
                    {
                        int index = statisticsCode[i].IndexOf(":");
                        string type = statisticsCode[i].Substring(index + 1, 1);
                        string code = statisticsCode[i].Substring(0, index);
                        if (type == "F")
                        {
                            statislist.Add(CodeNum.PdounInCode((KeyenceMcNet.ReadFloat(code).Content * 100).ToString()));
                        }
                        else
                        {
                            statislist.Add(CodeNum.StatisCodehu(KeyenceMcNet.ReadInt32(code).Content.ToString(), type));
                        }
                    }
                }
            }
            return statislist;

        }

        /// <summary>
        /// 触发发送设备故障信息
        /// </summary>
        /// <param name="status"></param>
        public void Namestation(string status)
        {
            Invoke(new Action(() =>
            {
                if (IsStart)
                {
                    if (faulttime != null)
                    {
                        for (int i = 0; i < faulttime.Length; i++)
                        {
                            if (faultsTable.Rows.Count > i)
                            {
                                ///编号
                                string faultID = faultsTable.Rows[i]["编号"].ToString();
                                if (faulttime[i] == 1)
                                {
                                    if (!faultsMap.ContainsKey(faultID))
                                    {
                                        string statssName = tbxStationName.Text; //
                                        string faultas = "1+" + statssName + "+" + textBox3.Text + "+" + status + "+"
                                         + faultsTable.Rows[i]["故障描述"] + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                        ShowMsg(faultas);
                                        Send(faultas);
                                        faultsMap.Add(faultID, faultas);
                                    }
                                }
                                else if (faulttime[i] == 0)
                                {
                                    string faulvor = "";
                                    if (faultsMap.TryGetValue(faultID, out faulvor))
                                    {
                                        if (!string.IsNullOrWhiteSpace(faulvor))
                                        {
                                            string shofault = faulvor + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                            ShowMsg(shofault);
                                            Send(shofault);
                                            faultsMap.Remove(faultID);
                                        }
                                    }
                                }
                                string jsonStr = JsonConvert.SerializeObject(faultsMap);
                                File.WriteAllText(pathText, jsonStr);
                            }
                        }
                    }
                }
            }));
        }

        /// <summary>
        /// 结束故障结束时候触发
        /// </summary>
        public void Fouddinog()
        {
            Invoke(new Action(() =>
            {
                if (IsStart)
                {
                    if (faultsMap.Count > 0 & faultsMap != null)
                    {
                        try
                        {
                            File.WriteAllText(pathText, "{}");
                            foreach (var item in faultsMap)
                            {
                                //输出
                                //1+故障所在工位+机台名称+设备状态+
                                //故障的描述+触发故障的结束时间 
                                string shofault = item.Value + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                ShowMsg(shofault);
                                Send(shofault);
                                faultsMap.Remove(item.Key);
                                if (faultsMap.Count == 0)
                                {
                                    break;
                                }
                            }
                        }
                        catch (Exception)
                        {

                        }
                    }
                }
            }));
        }

        Socket socketClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);

        /// <summary>
        /// 系统设置页面 > “保存”按钮的事件处理器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button25_Click(object sender, EventArgs e)
        {
            SYS_Model_Write();
        }

        /// <summary>
        /// 配方设置页面 > “保存”按钮的事件处理器
        /// </summary>
        private void button29_Click(object sender, EventArgs e)
        {
            SYS_Model_Write();
        }
        private void Button11_Click(object sender, EventArgs e)
        {
            // 触发通讯读
            //Button20_Click(null, null);
            Thread.Sleep(5);
            Application.DoEvents();
        }

        #endregion

        #region --------------- 显示列表 ---------------

        /// <summary>
        /// 显示列表表头
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertTable(object sender, EventArgs e)
        {
            //dataGridViewDynamic2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewDynamic2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            int countH = 0;
            countH = CodeNum.Tresultstrhomd(maxValue) + CodeNum.Tresultstrhomd(minValue) + CodeNum.Tresultstrhomd(testResult);

            //添加三列
            for (int i = 0; i < 7; i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic2.Columns.Add(new DataGridViewTextBoxColumn());
                dataGridViewDynamic2.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度

            }
            for (int i = 7; i < ((testItems.Length + countH) + 7); i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic2.Columns.Add(new DataGridViewTextBoxColumn());
                dataGridViewDynamic2.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度
            }
            dataGridViewDynamic2.Columns[0].HeaderText = "序号";
            dataGridViewDynamic2.Columns[1].HeaderText = "产品条码";
            dataGridViewDynamic2.Columns[2].HeaderText = "产品结果";
            dataGridViewDynamic2.Columns[3].HeaderText = "产品编号";
            dataGridViewDynamic2.Columns[4].HeaderText = "操作员";
            dataGridViewDynamic2.Columns[5].HeaderText = "上传状态";
            dataGridViewDynamic2.Columns[6].HeaderText = "测试时间";
            int a = 7;
            if (actualValue.Length > 0)
            {
                for (int i = 0; i < testItems.Length; i++)
                {
                    dataGridViewDynamic2.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                    dataGridViewDynamic2.Columns[a].HeaderText = testItems[i] + unitName[i];
                    a = a + 1;
                    if (!maxValue[i].Equals("NO"))
                    {
                        dataGridViewDynamic2.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                        dataGridViewDynamic2.Columns[a].HeaderText = testItems[i] + "上限" + unitName[i];
                        a = a + 1;
                    }
                    if (!minValue[i].Equals("NO"))
                    {
                        dataGridViewDynamic2.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                        dataGridViewDynamic2.Columns[a].HeaderText = testItems[i] + "下限" + unitName[i];
                        a = a + 1;
                    }
                    if (!testResult[i].Equals("NO"))
                    {
                        dataGridViewDynamic2.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                        dataGridViewDynamic2.Columns[a].HeaderText = testItems[i] + "结果";
                        a = a + 1;
                    }
                }
            }

            //dataGridViewDynamic2.Columns[3].HeaderText = "卷簧扭力";
            //dataGridViewDynamic2.Columns[4].HeaderText = "压力";
            //dataGridViewDynamic2.Columns[5].HeaderText = "行程";
            //dataGridViewDynamic2.Columns[8].HeaderText = "软件版本";
            //dataGridViewDynamic2.Columns[9].HeaderText = "文件版本";
            //dataGridViewDynamic2.Columns[10].HeaderText = "产品编码";
            //dataGridViewDynamic2.Columns[11].HeaderText = "产品名称";
            //dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;

        }//表格列表表头，根据设备不同自行增减 参照richTextBox.AppendText里边的字符串个数
         //表头不分左右！！！

        //显示行数据↓
        private void 显示结果_Left()
        {
            Invoke(new Action(() =>
            {
                if (dataGridViewDynamic2.RowCount > 5000)
                {
                    this.dataGridViewDynamic2.Rows.Clear();
                }
                var upState = "失败";
                if (Parameter_txt[2006] == "1")
                {
                    upState = "成功";
                }
                else if (OffLineType == 1)
                {
                    upState = "本地";
                }
                DateTime now = DateTime.Now;
                //添加行
                //int index = this.dataGridViewDynamic2.Rows.Add();
                DataGridViewRow dataGdVwRow = new DataGridViewRow();
                dataGdVwRow.CreateCells(this.dataGridViewDynamic2);
                dataGdVwRow.Cells[0].Value = Num;
                dataGdVwRow.Cells[1].Value = barcodeData;
                dataGdVwRow.Cells[2].Value = Value[9999];
                dataGdVwRow.Cells[3].Value = tbxProductName.Text;
                dataGdVwRow.Cells[4].Value = LoginUser.ToString();
                dataGdVwRow.Cells[5].Value = upState;
                dataGdVwRow.Cells[6].Value = now.ToString("MM-dd HH:mm:ss");//DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日" + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString();
                int a = 7;
                if (list.Count > 0)
                {

                    for (int i = 0; i < list.Count; i++)
                    {
                        dataGdVwRow.Cells[a].Value = list[i];
                        a = a + 1;
                        if (maxlist.Count > i)
                        {
                            if (!maxValue[i].Equals("NO"))
                            {
                                dataGdVwRow.Cells[a].Value = maxlist[i];
                                a = a + 1;
                            }
                        }

                        if (minlist.Count > i)
                        {
                            if (!minValue[i].Equals("NO"))
                            {
                                dataGdVwRow.Cells[a].Value = minlist[i];
                                a = a + 1;
                            }
                        }

                        if (resultlist.Count > i)
                        {
                            if (!testResult[i].Equals("NO"))
                            {
                                dataGdVwRow.Cells[a].Value = resultlist[i];
                                a = a + 1;
                            }
                        }
                    }
                }
                this.dataGridViewDynamic2.Rows.Insert(0, dataGdVwRow);
            }));
            //this.dataGridViewDynamic2.Rows.Insert(index, dataGridViewDynamic2.Rows[index]);
            ////每台机定制参数
            //this.dataGridViewDynamic2.Rows[index].Cells[2].Value = Parameter_txt[3692];
            //this.dataGridViewDynamic2.Rows[index].Cells[3].Value = Parameter_txt[3694];
            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[3696];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[3698];
            ////每台机定制参数

            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[1032];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[1034];
            //this.dataGridViewDynamic2.Rows[index].Cells[6].Value = Value[1036];
            //this.dataGridViewDynamic2.Rows[index].Cells[7].Value = Parameter_Model[17000];

            //this.dataGridViewDynamic2.Rows[index].Cells[8].Value = Parameter_Model[17500];
            //this.dataGridViewDynamic2.Rows[index].Cells[9].Value = Parameter_Model[17800];
            //this.dataGridViewDynamic2.Rows[index].Cells[10].Value = Parameter_Model[18000];
            //this.dataGridViewDynamic2.Rows[index].Cells[11].Value = Parameter_Model[18500];
            // dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;
        }
        private void InsertTable1(object sender, EventArgs e)
        {
            dataGridViewDynamic3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            int countH = 0;
            countH = CodeNum.WorkIDtrhomd(testItems, sequenceNum, "1") + CodeNum.WorkIDtrhomd(maxValue, sequenceNum, "1") + CodeNum.WorkIDtrhomd(minValue, sequenceNum, "1") + CodeNum.WorkIDtrhomd(testResult, sequenceNum, "1");

            //添加三列
            for (int i = 0; i < 7; i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic3.Columns.Add(new DataGridViewTextBoxColumn());
                dataGridViewDynamic3.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度
            }
            for (int i = 7; i < ((countH) + 7); i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic3.Columns.Add(new DataGridViewTextBoxColumn());
                // dataGridViewDynamic2.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度
            }
            dataGridViewDynamic3.Columns[0].HeaderText = "序号";
            dataGridViewDynamic3.Columns[1].HeaderText = "产品条码";
            dataGridViewDynamic3.Columns[2].HeaderText = "产品结果";
            dataGridViewDynamic3.Columns[3].HeaderText = "产品编号";
            dataGridViewDynamic3.Columns[4].HeaderText = "操作员";
            dataGridViewDynamic3.Columns[5].HeaderText = "上传状态";
            dataGridViewDynamic3.Columns[6].HeaderText = "测试时间";
            int a = 7;
            if (actualValue.Length > 0)
            {
                for (int i = 0; i < testItems.Length; i++)
                {
                    if (sequenceNum[i].Equals("1"))
                    {
                        dataGridViewDynamic3.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                        dataGridViewDynamic3.Columns[a].HeaderText = testItems[i] + unitName[i];
                        a = a + 1;
                        if (!maxValue[i].Equals("NO"))
                        {
                            dataGridViewDynamic3.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic3.Columns[a].HeaderText = testItems[i] + "上限" + unitName[i];
                            a = a + 1;
                        }
                        if (!minValue[i].Equals("NO"))
                        {
                            dataGridViewDynamic3.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic3.Columns[a].HeaderText = testItems[i] + "下限" + unitName[i];
                            a = a + 1;
                        }
                        if (!testResult[i].Equals("NO"))
                        {
                            dataGridViewDynamic3.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic3.Columns[a].HeaderText = testItems[i] + "结果";
                            a = a + 1;
                        }
                    }
                }
            }

            //dataGridViewDynamic2.Columns[3].HeaderText = "卷簧扭力";
            //dataGridViewDynamic2.Columns[4].HeaderText = "压力";
            //dataGridViewDynamic2.Columns[5].HeaderText = "行程";
            //dataGridViewDynamic2.Columns[8].HeaderText = "软件版本";
            //dataGridViewDynamic2.Columns[9].HeaderText = "文件版本";
            //dataGridViewDynamic2.Columns[10].HeaderText = "产品编码";
            //dataGridViewDynamic2.Columns[11].HeaderText = "产品名称";
            //dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;

        }//表格列表表头，根据设备不同自行增减 参照richTextBox.AppendText里边的字符串个数
         //表头不分左右！！！

        //显示行数据↓
        private void 显示结果_Left1(object sender, EventArgs e)
        {
            if (dataGridViewDynamic3.RowCount > 5000)
            {
                this.dataGridViewDynamic3.Rows.Clear();
            }
            var upState = "失败";
            if (Parameter_txt[2006] == "1")
            {
                upState = "成功";
            }
            else if (OffLineType == 1)
            {
                upState = "本地";
            }
            DateTime now = DateTime.Now;
            //添加行
            //int index = this.dataGridViewDynamic2.Rows.Add();
            DataGridViewRow dataGdVwRow = new DataGridViewRow();
            dataGdVwRow.CreateCells(this.dataGridViewDynamic3);
            dataGdVwRow.Cells[0].Value = Num;
            dataGdVwRow.Cells[1].Value = barcodeData;
            dataGdVwRow.Cells[2].Value = Value[9999];
            dataGdVwRow.Cells[3].Value = tbxProductName.Text;
            dataGdVwRow.Cells[4].Value = LoginUser.ToString();
            dataGdVwRow.Cells[5].Value = upState;
            dataGdVwRow.Cells[6].Value = now.ToString("yyyy年MM月dd日 HH:mm:ss");//DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日" + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString();
            int a = 7;
            if (list.Count > 0)
            {

                for (int i = 0; i < list.Count; i++)
                {
                    dataGdVwRow.Cells[a].Value = list[i];
                    a = a + 1;
                    if (maxlist.Count > i)
                    {
                        if (!(maxlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = maxlist[i];
                            a = a + 1;
                        }
                    }

                    if (minlist.Count > i)
                    {
                        if (!(minlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = minlist[i];
                            a = a + 1;
                        }
                    }

                    if (resultlist.Count > i)
                    {
                        if (!(resultlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = resultlist[i];
                            a = a + 1;
                        }
                    }
                }
            }
            this.dataGridViewDynamic3.Rows.Insert(0, dataGdVwRow);
            //this.dataGridViewDynamic2.Rows.Insert(index, dataGridViewDynamic2.Rows[index]);
            ////每台机定制参数
            //this.dataGridViewDynamic2.Rows[index].Cells[2].Value = Parameter_txt[3692];
            //this.dataGridViewDynamic2.Rows[index].Cells[3].Value = Parameter_txt[3694];
            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[3696];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[3698];
            ////每台机定制参数

            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[1032];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[1034];
            //this.dataGridViewDynamic2.Rows[index].Cells[6].Value = Value[1036];
            //this.dataGridViewDynamic2.Rows[index].Cells[7].Value = Parameter_Model[17000];

            //this.dataGridViewDynamic2.Rows[index].Cells[8].Value = Parameter_Model[17500];
            //this.dataGridViewDynamic2.Rows[index].Cells[9].Value = Parameter_Model[17800];
            //this.dataGridViewDynamic2.Rows[index].Cells[10].Value = Parameter_Model[18000];
            //this.dataGridViewDynamic2.Rows[index].Cells[11].Value = Parameter_Model[18500];
            // dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;
        }
        private void InsertTable2(object sender, EventArgs e)
        {
            dataGridViewDynamic4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            int countH = 0;
            countH = CodeNum.WorkIDtrhomd(testItems, sequenceNum, "2") + CodeNum.WorkIDtrhomd(maxValue, sequenceNum, "2") + CodeNum.WorkIDtrhomd(minValue, sequenceNum, "2") + CodeNum.WorkIDtrhomd(testResult, sequenceNum, "2");

            //添加三列
            for (int i = 0; i < 7; i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic4.Columns.Add(new DataGridViewTextBoxColumn());
                dataGridViewDynamic4.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度
            }
            for (int i = 7; i < ((countH) + 7); i++)//这里一定要改成对应个数，比如0~16是17个！
            {
                dataGridViewDynamic4.Columns.Add(new DataGridViewTextBoxColumn());
                //dataGridViewDynamic2.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; //设置所有列自适应宽度
            }
            dataGridViewDynamic4.Columns[0].HeaderText = "序号";
            dataGridViewDynamic4.Columns[1].HeaderText = "产品条码";
            dataGridViewDynamic4.Columns[2].HeaderText = "产品结果";
            dataGridViewDynamic4.Columns[3].HeaderText = "产品编号";
            dataGridViewDynamic4.Columns[4].HeaderText = "操作员";
            dataGridViewDynamic4.Columns[5].HeaderText = "上传状态";
            dataGridViewDynamic4.Columns[6].HeaderText = "测试时间";
            int a = 7;
            if (actualValue.Length > 0)
            {
                for (int i = 0; i < testItems.Length; i++)
                {
                    if (sequenceNum[i].Equals("2"))
                    {
                        dataGridViewDynamic4.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                        dataGridViewDynamic4.Columns[a].HeaderText = testItems[i] + unitName[i];
                        a = a + 1;
                        if (!maxValue[i].Equals("NO"))
                        {
                            dataGridViewDynamic4.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic4.Columns[a].HeaderText = testItems[i] + "上限" + unitName[i];
                            a = a + 1;
                        }
                        if (!minValue[i].Equals("NO"))
                        {
                            dataGridViewDynamic4.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic4.Columns[a].HeaderText = testItems[i] + "下限" + unitName[i];
                            a = a + 1;
                        }
                        if (!testResult[i].Equals("NO"))
                        {
                            dataGridViewDynamic4.Columns[a].Width = CodeNum.Doubtowule(textBox5.Text);
                            dataGridViewDynamic4.Columns[a].HeaderText = testItems[i] + "结果";
                            a = a + 1;
                        }
                    }
                }
            }

            //dataGridViewDynamic2.Columns[3].HeaderText = "卷簧扭力";
            //dataGridViewDynamic2.Columns[4].HeaderText = "压力";
            //dataGridViewDynamic2.Columns[5].HeaderText = "行程";
            //dataGridViewDynamic2.Columns[8].HeaderText = "软件版本";
            //dataGridViewDynamic2.Columns[9].HeaderText = "文件版本";
            //dataGridViewDynamic2.Columns[10].HeaderText = "产品编码";
            //dataGridViewDynamic2.Columns[11].HeaderText = "产品名称";
            //dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;

        }//表格列表表头，根据设备不同自行增减 参照richTextBox.AppendText里边的字符串个数
         //表头不分左右！！！

        //显示行数据↓
        private void 显示结果_Left2(object sender, EventArgs e)
        {
            if (dataGridViewDynamic4.RowCount > 5000)
            {
                this.dataGridViewDynamic4.Rows.Clear();
            }
            var upState = "失败";
            if (Parameter_txt[2006] == "1")
            {
                upState = "成功";
            }
            else if (OffLineType == 1)
            {
                upState = "本地";
            }
            DateTime now = DateTime.Now;
            //添加行
            //int index = this.dataGridViewDynamic2.Rows.Add();
            DataGridViewRow dataGdVwRow = new DataGridViewRow();
            dataGdVwRow.CreateCells(this.dataGridViewDynamic4);
            dataGdVwRow.Cells[0].Value = Num;
            dataGdVwRow.Cells[1].Value = barcodeData;
            dataGdVwRow.Cells[2].Value = Value[9999];
            dataGdVwRow.Cells[3].Value = tbxProductName.Text;
            dataGdVwRow.Cells[4].Value = LoginUser.ToString();
            dataGdVwRow.Cells[5].Value = upState;
            dataGdVwRow.Cells[6].Value = now.ToString("yyyy年MM月dd日 HH:mm:ss");//DateTime.Now.Year.ToString() + "年" + DateTime.Now.Month.ToString() + "月" + DateTime.Now.Day.ToString() + "日" + DateTime.Now.Hour.ToString() + ":" + DateTime.Now.Minute.ToString() + ":" + DateTime.Now.Second.ToString();
            int a = 7;
            if (list.Count > 0)
            {

                for (int i = 0; i < list.Count; i++)
                {
                    dataGdVwRow.Cells[a].Value = list[i];
                    a = a + 1;
                    if (maxlist.Count > i)
                    {
                        if (!(maxlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = maxlist[i];
                            a = a + 1;
                        }
                    }

                    if (minlist.Count > i)
                    {
                        if (!(minlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = minlist[i];
                            a = a + 1;
                        }
                    }

                    if (resultlist.Count > i)
                    {
                        if (!(resultlist[i]).Equals("null"))
                        {
                            dataGdVwRow.Cells[a].Value = resultlist[i];
                            a = a + 1;
                        }
                    }
                }
            }
            this.dataGridViewDynamic4.Rows.Insert(0, dataGdVwRow);
            //this.dataGridViewDynamic2.Rows.Insert(index, dataGridViewDynamic2.Rows[index]);
            ////每台机定制参数
            //this.dataGridViewDynamic2.Rows[index].Cells[2].Value = Parameter_txt[3692];
            //this.dataGridViewDynamic2.Rows[index].Cells[3].Value = Parameter_txt[3694];
            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[3696];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[3698];
            ////每台机定制参数

            //this.dataGridViewDynamic2.Rows[index].Cells[4].Value = Parameter_txt[1032];
            //this.dataGridViewDynamic2.Rows[index].Cells[5].Value = Parameter_txt[1034];
            //this.dataGridViewDynamic2.Rows[index].Cells[6].Value = Value[1036];
            //this.dataGridViewDynamic2.Rows[index].Cells[7].Value = Parameter_Model[17000];

            //this.dataGridViewDynamic2.Rows[index].Cells[8].Value = Parameter_Model[17500];
            //this.dataGridViewDynamic2.Rows[index].Cells[9].Value = Parameter_Model[17800];
            //this.dataGridViewDynamic2.Rows[index].Cells[10].Value = Parameter_Model[18000];
            //this.dataGridViewDynamic2.Rows[index].Cells[11].Value = Parameter_Model[18500];
            // dataGridViewDynamic2.FirstDisplayedScrollingRowIndex = dataGridViewDynamic2.Rows.Count - 1;
        }
        #endregion

        private async void saveInfo(string wroID)
        {
            await Task.Run(() =>
            {
                try
                {
                    LogMsg("生产数据本地保存中...");
                    DateTime times_Month = DateTime.Now;
                    string times_Month_string0 = times_Month.Year.ToString();
                    string times_Month_string1 = times_Month.Month.ToString();
                    Test(label2.Text + "\\" + times_Month_string0 + "年" + times_Month_string1 + "月" + "生产数据.mdb", wroID);
                    LogMsg("生产数据本地保存完成");
                }
                catch (Exception)
                {
                }
            });
        }

        private void Test(string conn, string wroID)
        {
            Invoke(new Action(() =>
            {
                mdbDatas.CreateAccessDatabase(conn);
                StringBuilder str = new StringBuilder("产品,当前工单号,工装编号,产品编码, 条码, 测试人,测试时间,测试结果");
                ArrayList arrayList = new ArrayList();
                object[] obj = new object[] { "产品", "当前工单号", "工装编号", "产品编码", "条码", "测试人", "测试时间", "测试结果" };
                arrayList.AddRange(obj);
                if (testItems.Length > 0)
                {
                    //object[] obj1 = new object[((testItems.Length)*4)];
                    for (int i = 0; i < (testItems.Length); i++)
                    {
                        arrayList.Add(testItems[i]);
                        if (!maxValue[i].Equals("NO") && !minValue[i].Equals("NO") && !testResult[i].Equals("NO"))
                        {
                            arrayList.Add(testItems[i] + "上限");
                            arrayList.Add(testItems[i] + "下限");
                            arrayList.Add(testItems[i] + "结果");
                        }
                    }

                    for (int i = 0; i < (testItems.Length); i++)
                    {
                        if (sequenceNum[i].Equals(wroID) || wroID.Equals("3"))
                        {
                            str.Append(",");
                            str.Append(testItems[i]);
                            if (!maxValue[i].Equals("NO") && !minValue[i].Equals("NO") && !testResult[i].Equals("NO"))
                            {
                                str.Append(",");
                                str.Append(testItems[i] + "上限");
                                str.Append(",");
                                str.Append(testItems[i] + "下限");
                                str.Append(",");
                                str.Append(testItems[i] + "结果");
                            }
                        }
                    }
                    // arrayList.AddRange(obj1);
                }

                mdbDatas.CreateMDBTable(conn, "Sheet1", arrayList);//new System.Collections.ArrayList(new object[] { "产品", "条码", "测试人", "测试时间","测试结果", "PLC配方", "文件版本","软件版本"}));
                mdb = new mdbDatas(conn);

                DateTime now = DateTime.Now;
                string CP = tbxProductName.Text == "" ? "无" : tbxProductName.Text;
                string barcode = barcodeData == "" ? " " : barcodeData;
                StringBuilder str1 = new StringBuilder();
                str1.Append("'" + CP + "',");
                str1.Append("'" + textBox6.Text + "',");
                str1.Append("'" + textBox41.Text + "',");
                str1.Append("'" + CodeNum.CodeStrfror(comboBox2.Text, codesDataM) + "',");
                str1.Append("'" + barcode + "',");
                str1.Append("'" + LoginUser.ToString() + "',");
                str1.Append("'" + now.ToString("yyyy年MM月dd日 HH:mm:ss") + "',");
                str1.Append("'" + Value[9999] + "'");

                if (list.Count > 0)
                {
                    for (int i = 0; i < list.Count; i++)
                    {

                        str1.Append(",");
                        string rst = " ";
                        if (list[i].Length > 0)
                        {
                            rst = list[i];
                        }
                        str1.Append("'" + rst + "'");
                        if (!maxlist[i].Equals("null") && !minlist[i].Equals("null"))
                        {
                            str1.Append(",");
                            str1.Append("'" + CodeNum.NullECoshu(maxlist[i]) + "'");
                            str1.Append(",");
                            str1.Append("'" + CodeNum.NullECoshu(minlist[i]) + "'");
                            str1.Append(",");
                            str1.Append("'" + CodeNum.NullECoshu(resultlist[i]) + "'");
                        }
                    }
                }

                string sql = "insert into Sheet1 (" + str + ") values (" + str1 + ")";


                bool result = mdb.Add(sql.ToString());
                //DataTable dt = new DataTable("Sheet1");
                //DataColumn pcode = new DataColumn("产品", typeof(string));
                //dt.Columns.Add(pcode);
                //DataColumn dccode = new DataColumn("条码", typeof(string));
                //dt.Columns.Add(dccode);
                //DataColumn dcName = new DataColumn("测试人", typeof(string));
                //dt.Columns.Add(dcName);
                //DataColumn dctime = new DataColumn("测试时间", typeof(string));
                //dt.Columns.Add(dctime);
                //DataColumn dcOKNG = new DataColumn("测试结果", typeof(string));
                //dt.Columns.Add(dcOKNG);


                //for (int i = 0; i < testItems.Length; i++)
                //{
                //    //var aa = i;
                //    DataColumn "aa"+i = new DataColumn(testItems[i], typeof(string));
                //    dt.Columns.Add(aa);
                //}
                ///////每台机定制参数
                ////DataColumn JHQS = new DataColumn("卷簧圈数", typeof(string));
                ////dt.Columns.Add(JHQS);
                ////DataColumn JHNL = new DataColumn("卷簧扭力", typeof(string));
                ////dt.Columns.Add(JHNL);
                ////DataColumn YL = new DataColumn("压力", typeof(string));
                ////dt.Columns.Add(YL);
                ////DataColumn XC = new DataColumn("行程", typeof(string));
                ////dt.Columns.Add(XC);
                ///////每台机定制参数

                //DataColumn dcheightPLC = new DataColumn("PLC配方", typeof(string));
                //dt.Columns.Add(dcheightPLC);
                //DataColumn dcFileVersion = new DataColumn("文件版本", typeof(string));
                //dt.Columns.Add(dcFileVersion);
                //DataColumn dcSoftwareVersion  = new DataColumn("软件版本", typeof(string));
                //dt.Columns.Add(dcSoftwareVersion);
                //DataRow dr = dt.NewRow();   
                //dt.Rows.Add(dr);
                ////dt.Rows.Add();
                //// dt.Rows.Add("时间", DateTime.Now);
                ////   dt.Rows.Add(dt.Rows[0].ItemArray);
                //DateTime now = DateTime.Now;
                //dr[0] = textBox2.Text;
                //dr[1] = barcodeID;
                //dr[2] = label41.Text;
                //dr[3] = now.ToString("yyyy年MM月dd日 HH:mm:ss");//times_Month_string0+"年"+ times_Month_string1+ "月"+ times_Month_string2 + "日"+" "+ times_Month_string3+ ":" + times_Month_string4 + ":" + times_Month_string5;
                //dr[4] = Value[9999];

                ////每台机定制参数
                //dr[5] = Parameter_txt[3692];
                //dr[6] = Parameter_txt[3694];
                //dr[7] = Parameter_txt[3696];
                //dr[8] = Parameter_txt[3698];
                ////每台机定制参数

                //dr[9] = Parameter_Model[17000]; 
                //dr[10] = Parameter_Model[17500];
                //dr[11] = Parameter_Model[17800];
                //mdb.DatatableToMdb("Sheet1", dt);

                mdb.CloseConnection();
            }));
        }

        private void saveSystemInfo(string conn, SytemInfoEntity systemInfo)
        {
            mdb = new mdbDatas(conn);
            DataTable table1 = mdb.Find("select * from SytemInfo where ID = '1'");

            if (table1.Rows.Count > 0)
            {

                DataRow row = table1.Rows[0];
                List<string> changeDetails = new List<string>();

                // 检查每个字段的变化
                if (row["IP"].ToString() != systemInfo.IP)
                {
                    changeDetails.Add($"IP：{row["IP"]} -> {systemInfo.IP}");
                }
                if (row["Port"].ToString() != systemInfo.Port)
                {
                    changeDetails.Add($"端口：{row["Port"]} -> {systemInfo.Port}");
                }
                if (row["Timeout"].ToString() != systemInfo.Timeout)
                {
                    changeDetails.Add($"连接超时：{row["Timeout"]} -> {systemInfo.Timeout}");
                }
                if (row["Url"].ToString() != systemInfo.Url)
                {
                    changeDetails.Add($"URL：{row["Url"]} -> {systemInfo.Url}");
                }
                if (row["Site"].ToString() != systemInfo.Site)
                {
                    changeDetails.Add($"站点：{row["Site"]} -> {systemInfo.Site}");
                }
                if (row["Resource"].ToString() != systemInfo.Resource)
                {
                    changeDetails.Add($"工位：{row["Resource"]} -> {systemInfo.Resource}");
                }
                if (row["Opration"].ToString() != systemInfo.Opration)
                {
                    changeDetails.Add($"工序：{row["Opration"]} -> {systemInfo.Opration}");
                }
                if (row["NcCode"].ToString() != systemInfo.NcCode)
                {
                    changeDetails.Add($"不合格代码：{row["NcCode"]} -> {systemInfo.NcCode}");
                }
                if (row["Password"].ToString() != systemInfo.Password)
                {
                    changeDetails.Add($"密码：{row["Password"]} -> {systemInfo.Password}");
                }
                if (row["User"].ToString() != systemInfo.User)
                {
                    changeDetails.Add($"用户：{row["User"]} -> {systemInfo.User}");
                }

                if (changeDetails.Count > 0)
                {
                    string sql = $"update [SytemInfo] set [IP] = '{systemInfo.IP}', [Port] = '{systemInfo.Port}', [Timeout] = '{systemInfo.Timeout}', " +
                        $"[NcCode] = '{systemInfo.NcCode}', [Opration] = '{systemInfo.Opration}', [Resource] = '{systemInfo.Resource}', " +
                        $"[Site] = '{systemInfo.Site}', [Url] = '{systemInfo.Url}', [User] = '{systemInfo.User}', [Password] = '{systemInfo.Password}' where [ID] = '1'";

                    var result = mdb.Change(sql);
                    if (result)
                    {
                        string changeLog = string.Join("\n", changeDetails);
                        loggerConfig.Trace($"【MES参数修改成功】\n修改详情：\n{changeLog}");
                        MessageBox.Show("保存成功");
                    }
                }
                else
                {

                    MessageBox.Show("没有任何变化，无需保存。");

                }
            }
            else
            {
                mdbDatas.CreateAccessDatabase(conn);
                mdbDatas.CreateMDBTable(conn, "SytemInfo", new System.Collections.ArrayList(new object[] { "ID", "IP", "Port", "Timeout", "NcCode", "Opration", "Password", "Resource", "Site", "Url", "User", "FileVersion", "SoftwareVersion", "UserCheckCmd", "UserCheckPass", "UserCheckFail", "CodeCheckCmd", "CodeCheckPass", "CodeSendCmd" }));
                DataTable dt = new DataTable("SytemInfo");
                DataColumn ID = new DataColumn("ID", typeof(string));
                dt.Columns.Add(ID);
                DataColumn IP = new DataColumn("IP", typeof(string));
                dt.Columns.Add(IP);
                DataColumn Port = new DataColumn("Port", typeof(string));
                dt.Columns.Add(Port);
                DataColumn Timeout = new DataColumn("Timeout", typeof(string));
                dt.Columns.Add(Timeout);
                DataColumn NcCode = new DataColumn("NcCode", typeof(string));
                dt.Columns.Add(NcCode);
                DataColumn Opration = new DataColumn("Opration", typeof(string));
                dt.Columns.Add(Opration);
                DataColumn Password = new DataColumn("Password", typeof(string));
                dt.Columns.Add(Password);
                DataColumn Resource = new DataColumn("Resource", typeof(string));
                dt.Columns.Add(Resource);
                DataColumn Site = new DataColumn("Site", typeof(string));
                dt.Columns.Add(Site);
                DataColumn Url = new DataColumn("Url", typeof(string));
                dt.Columns.Add(Url);
                DataColumn User = new DataColumn("User", typeof(string));
                dt.Columns.Add(User);
                DataColumn FileVersion = new DataColumn("FileVersion", typeof(string));
                dt.Columns.Add(FileVersion);
                DataColumn SoftwareVersion = new DataColumn("SoftwareVersion", typeof(string));
                dt.Columns.Add(SoftwareVersion);
                DataColumn UserCheckCmd = new DataColumn("UserCheckCmd", typeof(string));
                dt.Columns.Add(UserCheckCmd);
                DataColumn UserCheckFail = new DataColumn("UserCheckFail", typeof(string));
                dt.Columns.Add(UserCheckFail);
                DataColumn UserCheckPass = new DataColumn("UserCheckPass", typeof(string));
                dt.Columns.Add(UserCheckPass);
                DataColumn CodeCheckCmd = new DataColumn("CodeCheckCmd", typeof(string));
                dt.Columns.Add(CodeCheckCmd);
                DataColumn CodeCheckPass = new DataColumn("CodeCheckPass", typeof(string));
                dt.Columns.Add(CodeCheckPass);
                DataColumn CodeSendCmd = new DataColumn("CodeSendCmd", typeof(string));
                dt.Columns.Add(CodeSendCmd);
                DataRow dr = dt.NewRow();
                dt.Rows.Add(dr);

                dr[0] = "1";
                dr[1] = systemInfo.IP;
                dr[2] = systemInfo.Port;
                dr[3] = systemInfo.Timeout;
                dr[4] = systemInfo.NcCode;
                dr[5] = systemInfo.Opration;
                dr[6] = "";
                dr[7] = systemInfo.Resource;
                dr[8] = systemInfo.Site;
                dr[9] = systemInfo.Url;
                dr[10] = "";
                dr[11] = ""; //systemInfo.FileVersion;
                dr[12] = ""; //systemInfo.SoftwareVersion;
                dr[13] = ""; //systemInfo.UserCheckCmd;
                dr[14] = ""; //systemInfo.UserCheckPass;
                dr[15] = ""; //systemInfo.UserCheckFail;
                dr[16] = ""; //systemInfo.CodeCheckCmd;
                dr[17] = ""; //systemInfo.CodeCheckPass;
                dr[18] = ""; //systemInfo.CodeSendCmd;
                mdb.DatatableToMdb("SytemInfo", dt);
            }

            mdb.CloseConnection();
        }

        /// <summary>
        /// ///读取条码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button19_Click(object sender, EventArgs e)
        {

            if (plcConn == true)
            {
                var Read_data = KeyenceMcNet.ReadInt32("D1000").Content;      // 读取寄存器int值 

                if (Read_data == 1)
                {
                    Invoke(new Action(() =>
                    {
                        lblScanBarcodeSatus.ForeColor = B;  // 扫码状态指示灯
                        lblValidationStatus.ForeColor = B;  // 验证状态指示灯
                        lblUploadStatus.ForeColor = B;      // 上传状态指示灯

                        lblRunningStatus.ForeColor = B;
                        lblProductResult.Text = resources.GetString("label_Value"); // 待机
                        lblProductResult.ForeColor = B;
                        lblProductResult.BackColor = W;

                        LogMsg("准备读条码....");
                        string rawBarcode = KeyenceMcNet.ReadString("D1100", 10).Content;
                        this.barcodeInfo = CodeNum.StrVBcd(rawBarcode);
                        txtShowBarcode.Text = barcodeInfo;
                        LogMsg($"条码【D1100】 = {barcodeInfo} ");

                        // 条码为空
                        if (string.IsNullOrEmpty(barcodeInfo))
                        {
                            lblScanBarcodeSatus.ForeColor = R;
                            txtShowBarcode.Text = resources.GetString("barCode_State"); // 未获取到条码，请重新扫描！
                            LogMsg("未获取到条码！");
                            try
                            {
                                KeyenceMcNet.Write("D1005", 1);
                                LogMsg("反馈条码验证【D1005】 = 1");
                            }
                            catch (Exception ex)
                            {
                                LogMsg(ex.ToString());
                            }
                            return;
                        }

                        if (chkBypassBarcodeValidation.Checked == false)
                        {
                            if (chkBypassFixtureVali.Checked == false)
                            {
                                string[] frockA = CodeNum.Confrock(comboBox2.Text);
                                string[] frockB = textBox41.Text.ToString().Split('+');
                                if (!CodeNum.CopareArr(frockA, frockB))
                                {
                                    if (frockA.Contains(barcodeInfo))
                                    {

                                        if (string.IsNullOrWhiteSpace(textBox41.Text))
                                        {
                                            textBox41.Text = barcodeInfo;
                                        }
                                        else
                                        {
                                            if (!frockB.Contains(barcodeInfo))
                                            {
                                                textBox41.Text = textBox41.Text + "+" + barcodeInfo;
                                            }
                                        }
                                        string[] frockC = textBox41.Text.ToString().Split('+');
                                        if (CodeNum.CopareArr(frockA, frockC))
                                        {
                                            mdb = new mdbDatas(path4);
                                            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");
                                            if (table1.Rows.Count > 0)
                                            {
                                                string sql = "update [SytemSet] set [BoardBeat]='" + textBox41.Text + "' where [ID] = '1'";
                                                var result = mdb.Change(sql);
                                            }
                                            mdb.CloseConnection();
                                            lblActionTips.Text = resources.GetString("scanning");
                                        }
                                        lblScanBarcodeSatus.ForeColor = G;
                                        lblRunningStatus.ForeColor = G;
                                        lblRunningStatus.Text = resources.GetString("Fixture_OK");
                                        try
                                        {
                                            KeyenceMcNet.Write("D1003", 2);
                                        }
                                        catch (Exception ex)
                                        { LogMsg(ex.ToString()); }
                                        LogMsg("反馈条码验证【D1003】 = 2");
                                        return;
                                    }
                                    else
                                    {
                                        lblRunningStatus.ForeColor = R;
                                        lblRunningStatus.Text = resources.GetString("Fixture_NG");
                                        lblValidationStatus.ForeColor = R;
                                        LogMsg("工装编号验证失败！");
                                        try
                                        {
                                            KeyenceMcNet.Write("D1005", 2);
                                            LogMsg("反馈条码验证【D1005】 = 2");
                                        }
                                        catch (Exception ex)
                                        { LogMsg(ex.ToString()); }
                                        return;
                                    }
                                }
                            }

                            if (barcodeInfo != "1")
                            {
                                DateTime times_Month = DateTime.Now;
                                string times_Month_string0 = times_Month.Year.ToString();
                                string times_Month_string1 = times_Month.Month.ToString();
                                string context = label2.Text + "\\" + times_Month_string0 + "年" + times_Month_string1 + "月" + "生产数据.mdb";
                                mdb = new mdbDatas(context);
                                bool boFbarcod = mdb.FarettuFind("select * from Sheet1 where 条码 = '" + barcodeInfo + "'");
                                mdb.CloseConnection();
                                if (boFbarcod)
                                {
                                    lblRunningStatus.ForeColor = R;
                                    lblRunningStatus.Text = resources.GetString("barCode_repeat");
                                    lblValidationStatus.ForeColor = R;
                                    LogMsg("条码重复扫过，请重新扫描！");
                                    try
                                    {
                                        KeyenceMcNet.Write("D1005", 1);
                                        LogMsg("反馈条码验证【D1005】 = 1");
                                    }
                                    catch (Exception ex)
                                    { LogMsg(ex.ToString()); }
                                    return;
                                }
                            }

                            string bardid = CodeNum.Condend(comboBox2.Text);

                            //验证条码
                            if (string.IsNullOrWhiteSpace(bardid))
                            {
                                lblRunningStatus.ForeColor = R;
                                lblRunningStatus.Text = resources.GetString("barCode_NG");
                                lblValidationStatus.ForeColor = R;
                                LogMsg("没有选择条码验证！！");
                                try
                                {
                                    KeyenceMcNet.Write("D1005", 1);
                                    LogMsg("反馈条码验证【D1005】 = 1");
                                }
                                catch (Exception ex)
                                { LogMsg(ex.ToString()); }
                                return;
                            }
                            else
                            {
                                int barlong = bardid.Length;
                                bool barboo = false;
                                if (barcodeInfo.Length > 15 && barcodeInfo.Length > barlong)
                                {
                                    string barsper = barcodeInfo.Substring(0, barlong);
                                    if (bardid.Equals(barsper))
                                    {
                                        barboo = true;
                                    }
                                }
                                if (barboo == false)
                                {
                                    lblRunningStatus.ForeColor = R;
                                    lblRunningStatus.Text = resources.GetString("barCode_NG");
                                    lblValidationStatus.ForeColor = R;
                                    LogMsg("没有写条码验证！！");
                                    try
                                    {
                                        KeyenceMcNet.Write("D1005", 1);
                                        LogMsg("反馈条码验证【D1005】 = 1");
                                    }
                                    catch (Exception ex)
                                    { LogMsg(ex.ToString()); }
                                    return;
                                }
                            }
                        }

                        //条码规则验证();
                        if (OffLineType == 1)
                        {
                            lblScanBarcodeSatus.ForeColor = G;
                            lblRunningStatus.ForeColor = G;
                            lblRunningStatus.Text = resources.GetString("barCode_OK");
                            try
                            {
                                KeyenceMcNet.Write("D1003", 1);
                            }
                            catch (Exception ex)
                            { LogMsg(ex.ToString()); }
                            LogMsg("反馈条码验证【D1003】 = 1");
                        }
                        else
                        {
                            if (checkBox1.Checked == true)
                            {
                                工单号 = textBox6.Text;
                                绑定工单(null, null);
                                //textBox6.Text = 工单号;
                            }
                            Button9_Click(null, null);//code条码验证
                            if (Parameter_txt[2002] == "1")
                            {

                                lblRunningStatus.ForeColor = G;
                                lblRunningStatus.Text = resources.GetString("barCode_OK");
                                lblValidationStatus.ForeColor = G;
                                try
                                {
                                    KeyenceMcNet.Write("D1003", 1);
                                }
                                catch (Exception ex)
                                { LogMsg(ex.ToString()); }
                                LogMsg("反馈条码验证【D1003】 = 1");
                            }
                            else if (Parameter_txt[2004] == "1")
                            {
                                lblRunningStatus.ForeColor = R;
                                lblRunningStatus.Text = resources.GetString("barCode_NG_MES");
                                lblValidationStatus.ForeColor = R;
                                try
                                {
                                    KeyenceMcNet.Write("D1005", 1);
                                }
                                catch (Exception ex)
                                { LogMsg(ex.ToString()); }
                                LogMsg("反馈条码验证【D1005】 = 1");
                            }
                            Thread.Sleep(200);
                            Application.DoEvents();
                        }

                        // busTcpClient.Write("2000", Convert.ToInt16(0));
                        LogMsg("条码读取完成.........");
                        lblActionTips.Text = resources.GetString("ScanBarCode_OK");
                    }));
                }
                if (Read_data == 2)
                {
                    Invoke(new Action(() =>
                    {
                        string barcodeID2 = KeyenceMcNet.ReadString("D1100", 10).Content;
                        string barcodeID1 = CodeNum.StrVBcd(barcodeID2);
                        txtShowBarcode.Text = barcodeID1;
                        LogMsg("条码【D1100】 = " + barcodeID1);
                        barcodeInfo = barcodeID1;
                        if (string.IsNullOrEmpty(barcodeInfo))
                        {
                            lblScanBarcodeSatus.ForeColor = R;
                            txtShowBarcode.Text = resources.GetString("barCode_State");
                            LogMsg("未获取到条码，请重新扫描！");
                            try
                            {
                                KeyenceMcNet.Write("D1005", 1);
                                LogMsg("反馈条码验证【D1005】 = 1");
                            }
                            catch (Exception ex)
                            { LogMsg(ex.ToString()); }
                            return;
                        }
                        string[] frockA = textBox42.Text.ToString().Split('+');
                        if (frockA.Contains(barcodeInfo))
                        {
                            lblScanBarcodeSatus.ForeColor = G;
                            lblRunningStatus.ForeColor = G;
                            lblRunningStatus.Text = resources.GetString("material_OK");
                            try
                            {
                                KeyenceMcNet.Write("D1003", 3);
                            }
                            catch (Exception ex)
                            { LogMsg(ex.ToString()); }
                            LogMsg("反馈条码验证【D1003】 =3");
                            return;
                        }
                        else
                        {
                            lblRunningStatus.ForeColor = R;
                            lblRunningStatus.Text = resources.GetString("material_NG");
                            lblValidationStatus.ForeColor = R;
                            LogMsg("物料验证失败，请重新扫描！");
                            try
                            {
                                KeyenceMcNet.Write("D1005", 3);
                                LogMsg("反馈条码验证【D1005】 =3");
                            }
                            catch (Exception ex)
                            { LogMsg(ex.ToString()); }
                            return;
                        }
                    }));
                }
            }
        }

        public Label lb = new Label();
        Color G = Color.Green;
        Color R = Color.Red;
        Color W = Color.White;
        Color B = Color.Black;
        Color O = Color.Orange;
        List<string> AstrName = null;
        string barcodeData = "1";
        string barcodeRepat = "2";

        /// <summary>
        /// 获取生产结果
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button20_Click(object sender, EventArgs e, string wroID)
        {
            sw.Start();
            Console.WriteLine("Button20 Start");
            //D1820:I-0
            // int Read_data = busTcpClient.ReadInt16("2010").Content;  // 读取寄存器int值 
            if (plcConn == true)
            {
                //是否可读 1=可读
                var short_D100 = KeyenceMcNet.ReadInt32("D1200").Content;
                //LogMsg("生产结果【D1200】 = " + short_D100);
                Console.WriteLine("short_D100 = " + DateTime.Now.Hour.ToString() +
                DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" +
                DateTime.Now.Millisecond.ToString() + ":" + short_D100);

                if (short_D100 == 1)
                {
                    //  Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff:ffffff"));
                    bool isRepat = false;
                    Invoke(new Action(() =>
                    {
                        barcodeData = barcodeInfo;
                        Console.WriteLine("开始读取数据 = " + DateTime.Now.Hour.ToString() +
                        DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" +
                        DateTime.Now.Millisecond.ToString() + ":" + short_D100);
                        lblActionTips.Text = resources.GetString("begin_read_data");

                        LogMsg("生产结果数据读取中....");
                        string[] Read_string = new string[10000];
                        //string[] Read_string1 = new string[100];
                        AstrName = new List<string>();
                        list = new List<string>();
                        // beatlist = new List<string>();
                        maxlist = new List<string>();//
                        minlist = new List<string>();
                        resultlist = new List<string>();
                        // vulpalist = new List<string>();
                        //Read_string[1030] = KeyenceMcNet.ReadString("D1030",1).Content;//高度上限
                        //Read_string[1032] = KeyenceMcNet.ReadString("D1032", 1).Content;//高度下限
                        //Read_string[3686] = KeyenceMcNet.ReadString("D3686", 1).Content;//折叠尺寸


                        //判断是否最后读取产品条码
                        if (checkBox9.Checked == true)
                        {
                            //适用于转盘机台、最后在上传条码
                            string barcodeID2 = KeyenceMcNet.ReadString("D1150", 10).Content;
                            barcodeData = CodeNum.StrVBcd(barcodeID2);
                        }
                        if (checkBox11.Checked == true)
                        {
                            if (barcodeRepat.Equals(barcodeData))
                            {
                                try
                                {
                                    KeyenceMcNet.Write("D1202", 1);
                                }
                                catch (Exception ex)
                                { LogMsg(ex.ToString()); }
                                LogMsg("未上传数据直接反馈生产结果读取反馈【D1202】 = 1");
                                Thread.Sleep(200);
                                Application.DoEvents();
                                lblRunningStatus.Text = "数据条码重复上传";
                                lblRunningStatus.ForeColor = Color.Red;
                                isRepat = true;
                            }
                            barcodeRepat = barcodeData;
                        }
                        //Read_string[3676] = KeyenceMcNet.ReadInt32("D1080").Content.ToString();//生产总数
                        //Read_string[3678] = KeyenceMcNet.ReadInt32("D1084").Content.ToString();//工单数量
                        //Read_string[3680] = KeyenceMcNet.ReadInt32("D1086").Content.ToString();//完成数量
                        //Read_string[3682] = KeyenceMcNet.ReadInt32("D1088").Content.ToString();//合格数量
                        //Read_string[3684] = KeyenceMcNet.ReadInt32("D1090").Content.ToString();//生产节拍
                        //Read_string[3686] = KeyenceMcNet.ReadInt32("D1082").Content.ToString();//保养计数
                        //Read_string[3690] = KeyenceMcNet.ReadInt32("D1076").Content.ToString();//NG数量
                        //D1820:I-0
                        //实际测试项的值

                        // actualValue = userCollection.AsEnumerable().Select(row => row["BoardCode"].ToString()).ToArray();
                        //   maxValue = userCollection.AsEnumerable().Select(row => row["MaxBoardCode"].ToString()).ToArray();
                        //  minValue = userCollection.AsEnumerable().Select(row => row["MinBoardCode"].ToString()).ToArray();
                        //  beat = userCollection.AsEnumerable().Select(row => row["BeatBoardCode"].ToString()).ToArray();
                        //  testResult = userCollection.AsEnumerable().Select(row => row["ResultBoardCode"].ToString()).ToArray();
                        if (boardTable.Rows.Count > 0)
                        {
                            for (int i = 0; i < boardTable.Rows.Count; i++)
                            {
                                if (sequenceNum[i].Equals(wroID) || wroID.Equals("3"))
                                {
                                    AstrName.Add(testItems[i]);
                                    list.Add(PLCCodeNum(boardTable.Rows[i]["BoardCode"].ToString()));
                                    maxlist.Add(PLCCodeNum(boardTable.Rows[i]["MaxBoardCode"].ToString()));
                                    minlist.Add(PLCCodeNum(boardTable.Rows[i]["MinBoardCode"].ToString()));
                                    resultlist.Add(PLCCodeNum(boardTable.Rows[i]["ResultBoardCode"].ToString()));
                                    //    beatlist.Add(PLCCodeNum(boardTable.Rows[i]["BeatBoardCode"].ToString()));
                                }
                            }
                        }

                        Read_string[3688] = KeyenceMcNet.ReadInt32("D1078").Content.ToString();//产品状态
                        Parameter_txt[3688] = Convert.ToString(Convert.ToDouble(Read_string[3688]));

                        // Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff:ffffff"));
                        ////每台机定制参数
                        //Read_string[3692] = KeyenceMcNet.ReadInt32("D1092").Content.ToString();//卷簧圈数
                        //Read_string[3694] = KeyenceMcNet.ReadInt32("D1094").Content.ToString();//卷簧扭力
                        //Read_string[3696] = KeyenceMcNet.ReadInt32("D1096").Content.ToString();//压力
                        //Read_string[3698] = KeyenceMcNet.ReadInt32("D1098").Content.ToString();//行程
                        ////每台机定制参数

                        //Parameter_txt[1030] = Convert.ToString(Convert.ToDouble(Read_string[1030]) / 100);
                        //Parameter_txt[1032] = Convert.ToString(Convert.ToDouble(Read_string[1032]) / 100);
                        //Parameter_txt[3686] = Convert.ToString(Convert.ToDouble(Read_string[3686]));

                        //Parameter_txt[3676] = Convert.ToString(Convert.ToDouble(Read_string[3676]));
                        //Parameter_txt[3678] = Convert.ToString(Convert.ToDouble(Read_string[3678]));
                        //Parameter_txt[3680] = Convert.ToString(Convert.ToDouble(Read_string[3680]));
                        //Parameter_txt[3682] = Convert.ToString(Convert.ToDouble(Read_string[3682]));
                        //Parameter_txt[3684] = Convert.ToString(Convert.ToDouble(Read_string[3684]));
                        //Parameter_txt[3686] = Convert.ToString(Convert.ToDouble(Read_string[3686]));
                        //Parameter_txt[3690] = Convert.ToString(Convert.ToDouble(Read_string[3690]));

                        //每台机定制参数
                        //Parameter_txt[3692] = Convert.ToString(Convert.ToDouble(Read_string[3692]));
                        //Parameter_txt[3694] = Convert.ToString(Convert.ToDouble(Read_string[3694]));
                        //Parameter_txt[3696] = Convert.ToString(Convert.ToDouble(Read_string[3696]));
                        //Parameter_txt[3698] = Convert.ToString(Convert.ToDouble(Read_string[3698]));
                        //每台机定制参数
                        //for (int i = 0; i < list.Count; i++) 
                        //{
                        //    Parameter_txt1[i] = Convert.ToString(Convert.ToDouble(list[i]));
                        //}

                    }));

                    if (isRepat)
                    {
                        return;
                    }

                    Invoke(new Action(() =>
                    {
                        if (Parameter_txt[3688] == "3")
                        {
                            lblProductResult.Text = "OK";
                            lblProductResult.ForeColor = W;
                            lblProductResult.BackColor = G;
                            Value[9999] = "OK";

                        }
                        else
                        {
                            lblProductResult.Text = "NG";
                            lblProductResult.ForeColor = W;
                            lblProductResult.BackColor = R;
                            Value[9999] = "NG";
                        }
                    }));

                    // 上传数据
                    if (OffLineType == 0)
                    {
                        Invoke(new Action(() =>
                        {
                            lblRunningStatus.ForeColor = B;
                            lblRunningStatus.Text = resources.GetString("Mes_upload");
                            lblActionTips.Text = resources.GetString("Wait");
                            lblUploadStatus.ForeColor = O;
                            if (Value[9999] == "OK")
                            {
                                产品结果 = true;
                            }
                            else
                            {
                                产品结果 = false;
                            }
                            Button7_Click(null, null);
                            if (Parameter_txt[2006] == "1")
                            {
                                lblRunningStatus.Text = resources.GetString("Mes_upload_OK");
                                lblUploadStatus.ForeColor = G;
                                //KeyenceMcNet.Write("D3672", Convert.ToInt16(1));
                            }
                            else if (Parameter_txt[2008] == "1")
                            {
                                lblRunningStatus.Text = resources.GetString("Mes_upload_NG");
                                lblUploadStatus.ForeColor = R;
                                lblActionTips.Text = resources.GetString("Re_upload");
                                // KeyenceMcNet.Write("D3674", Convert.ToInt16(1));
                            }
                        }));
                    }
                    //生产数据上传看板
                    if (IsStart)
                    {
                        SendReceived();
                    }
                    //保存数据 

                    // Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff:ffffff"));
                    Invoke(new Action(() =>
                    {
                        if (wroID.Equals("3"))
                        {
                            显示结果_Left();
                        }
                        else if (wroID.Equals("1"))
                        {
                            显示结果_Left1(null, null);
                        }
                        else if (wroID.Equals("2"))
                        {
                            显示结果_Left2(null, null);
                        }
                        //显示结果_Left(null, null);
                        //if (Parameter_txt[3688] != "0")
                        //{
                        Num = Num + 1;

                        //label83.Text = Parameter_txt[1030]; label84.Text = Parameter_txt[1032]; //label.Text = Parameter_Model[17000];
                        //textBox6.Text = Parameter_txt[1034];
                        lblRunningStatus.Text = resources.GetString("Read_data");
                        lblActionTips.Text = resources.GetString("Wait");
                        saveInfo(wroID);
                        //saveInfo();
                        //Savecsv(null, null);
                        lblRunningStatus.Text = resources.GetString("Read_data_OK");
                        lblActionTips.Text = resources.GetString("Continue_production");
                        //textBox17.Text = Parameter_txt[3676];//生产总数
                        //                                     //textBox16.Text = Parameter_txt[3678];//工单数量 
                        //textBox9.Text = Parameter_txt[3678];//工单数量
                        //textBox15.Text = Parameter_txt[3680];//完成数量数量 
                        //textBox10.Text = Parameter_txt[3682];//合格数量

                        //textBox14.Text = Parameter_txt[3684];//生产节拍
                        //if (Parameter_txt[1088] != "0" && Parameter_txt[1080] != "0" && Parameter_txt[1080] != null)
                        //{
                        //    textBox13.Text = Convert.ToString(Math.Round(Convert.ToDouble(Parameter_txt[1088]) / Convert.ToDouble(Parameter_txt[1080]) * 100, 1)) + "%";//合格率
                        //}
                        //if (Parameter_txt[3684] != "0" && Parameter_txt[3684] != null)
                        //{
                        //    textBox14.Text = Convert.ToString(Math.Round(Convert.ToDouble(Parameter_txt[3684]) / Convert.ToDouble(10), 1)) + "s";//生产节拍
                        //}
                        //textBox7.Text = Parameter_txt[3686];   //保养计数
                        //textBox11.Text = Parameter_txt[3690];  //NG数量
                        LogMsg("数据读取完成");
                        LogMsg("生产结果读取完成...");
                        Console.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:fff:ffffff"));
                        try
                        {
                            KeyenceMcNet.Write("D1202", 1);
                        }
                        catch (Exception ex)
                        { LogMsg(ex.ToString()); }
                        LogMsg("生产结果读取反馈【D1202】 = 1");
                        Thread.Sleep(200);
                        Application.DoEvents();
                        lblRunningStatus.Text = resources.GetString("scanning");
                        //}

                    }));
                }
            }

            sw.Stop();
            Console.WriteLine("Button20 End");

            Console.WriteLine($"总耗时：{sw.Elapsed.TotalSeconds:F2}秒");
        }

        //static Stopwatch stopwatch = new Stopwatch();
        private static string PLCCodeNum(string plccode)
        {
            // stopwatch.Start();
            string redaa = "";
            if (plccode.Equals("NO"))
            {

                redaa = "null";
            }
            else if (plccode.Contains(":"))
            {
                int index = plccode.IndexOf(":");
                int index1 = plccode.IndexOf("-");
                string type = plccode.Substring(index + 1, 1);
                ushort len = 0;
                ushort.TryParse(plccode.Substring(index1 + 1, 1), out len);
                string code = plccode.Substring(0, index);

                if (type == "H")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    //  Console.WriteLine($"H开始线程执行时间: {stopwatch.ElapsedMilliseconds} 毫秒");
                    OperateResult<short> readss = KeyenceMcNet.ReadInt16(code);
                    //  Console.WriteLine($"H结束线程执行时间: {stopwatch.ElapsedMilliseconds} 毫秒");
                    if (readss.IsSuccess)
                    {
                        aa = TypeRead.Typerdess(readss.Content.ToString(), len.ToString());
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }       // short
                else if (type == "I")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    OperateResult<int> readss = KeyenceMcNet.ReadInt32(code);
                    if (readss.IsSuccess)
                    {
                        aa = TypeRead.Typerdess(readss.Content.ToString(), len.ToString());
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }  // Int32
                else if (type == "F")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    OperateResult<float> readss = KeyenceMcNet.ReadFloat(code);
                    if (readss.IsSuccess)
                    {
                        aa = TypeRead.Typerdess(readss.Content.ToString(), len.ToString());
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }  // float
                else if (type == "J")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    OperateResult<int> readss = KeyenceMcNet.ReadInt32(code);
                    if (readss.IsSuccess)
                    {

                        aa = CodeNum.PdounInCode(readss.Content.ToString());
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }  // Int32 
                else if (type == "N")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    OperateResult<int> readss = KeyenceMcNet.ReadInt32(code);
                    if (readss.IsSuccess)
                    {
                        aa = CodeNum.PNumOKAG(readss.Content.ToString());
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }  // Int32 
                else if (type == "O")
                {
                    string aa = "";
                    // var ss = KeyenceMcNet.ReadInt32(code).Content.ToString();
                    OperateResult<int> readss = KeyenceMcNet.ReadInt32(code);
                    if (readss.IsSuccess)
                    {
                        aa = readss.Content.ToString();
                    }
                    else
                    {
                        aa = "null";
                    }
                    redaa = aa;
                }  // Int32 
                else if (type == "S")
                {
                    string ss = "";
                    // var ss = KeyenceMcNet.ReadString(code, len).Content;
                    OperateResult<string> readss = KeyenceMcNet.ReadString(code, len);
                    if (readss.IsSuccess)
                    {
                        ss = CodeNum.StrVBcd(readss.Content.ToString());
                    }
                    else
                    {
                        ss = "null";
                    }
                    redaa = ss;
                }  // string
                else
                {
                    var ss = KeyenceMcNet.ReadString(code, len).Content;
                    redaa = ss;
                }                   // string
            }
            else
            {
                string aa = "";
                OperateResult<int> readss = KeyenceMcNet.ReadInt32(plccode);
                if (readss.IsSuccess)
                {

                    aa = CodeNum.PNumCode(readss.Content.ToString());
                }
                else
                {
                    aa = "null";
                }
                redaa = aa;
            }
            // 停止计时器
            //stopwatch.Stop();

            // 输出执行时间
            // Console.WriteLine($"线程执行时间: {stopwatch.ElapsedMilliseconds} 毫秒");
            return redaa;
        }

        /// <summary>
        /// 读取
        /// </summary>
        /// <param name="KeyenceMcNet"></param>
        /// <param name="strarr"></param>
        /// <param name="strtyp"></param>
        /// <returns></returns>
        public List<string> PCodenum(string[] strarr, string strtyp)
        {
            List<string> listarr = new List<string>();
            if (strarr.Length > 0)
            {
                string boardBeat = "   ";//
                for (int i = 0; i < strarr.Length; i++)
                {
                    string beatcode = strarr[i];//
                    if (beatcode.Contains("+"))
                    {
                        int count = 0;
                        string newStr = beatcode.Replace("+", "");
                        int.TryParse(newStr, out count);
                        for (int j = 0; j < count; j++)
                        {
                            listarr.Add(boardBeat);
                        }
                    }
                    else
                    {
                        if (beatcode.Equals("NO"))
                        {
                            boardBeat = "   ";
                            listarr.Add(boardBeat);
                        }
                        else
                        {
                            boardBeat = KeyenceMcNet.ReadInt32(beatcode).Content.ToString();
                            boardBeat = Bocmdb.BoardCode(boardBeat, strtyp);
                            listarr.Add(boardBeat);
                        }
                    }
                }
            }
            return listarr;
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button13_Click(object sender, EventArgs e)
        {
            //打开文件对话框，导出文件
            //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "保存文件";
            saveFileDialog1.Filter = "Excel文件(*.xls,*.xlsx,*.xlsm)|*.xls,*.xlsx,*.xlsm";
            saveFileDialog1.FileName = "历史数据.xls"; //设置默认另存为的名字

            string dataBaseUrl = textBoxPath.Text + ".mdb";
            mdb = new mdbDatas(dataBaseUrl);
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string txtPath = this.saveFileDialog1.FileName;
                DateTime times_Start = dateTimePicker1.Value;
                DateTime times_End = dateTimePicker2.Value;
                string times_Start_string = times_Start.ToString();
                string times_End_string = times_End.ToString();//dateTimePicker2.Value.ToString("yyyy/MM/dd").ToString() + " 23:59:59";
                DateTime times_Start_Sub = DateTime.Parse(times_Start_string);
                DateTime times_End_Sub = DateTime.Parse(times_End_string);
                if (textBoxPath.Text.Length < 1)
                {
                    MessageBox.Show("请先选择左侧数据源！");
                    return;
                }

                StringBuilder sql = new StringBuilder();

                sql.Append("select * from [Sheet1] where CDate(Format(测试时间,'yyyy/MM/dd HH:mm:ss')) between  #" + times_Start_Sub + "#" + " and #" + Convert.ToDateTime(times_End_Sub) + "# ");

                if (textBox_Code.Text != "")
                {
                    sql.Append("and 条码 = '" + textBox_Code.Text + "'");
                }
                if (textBox1.Text != "")
                {
                    sql.Append("and 产品 = '" + textBox1.Text + "' ");
                }

                DataTable table = mdb.Find(sql.ToString());

                NPOIHelper.DataTableToExcel(table, txtPath);
            }
            mdb.CloseConnection();
        }

        private void Button15_Click(object sender, EventArgs e)
        {
            SaveParameter_MES();
            LoadParameter_MES();
            Button8_Click(null, null);
        }

        private void Button17_Click(object sender, EventArgs e)
        {
            //条码规则验证();
        }

        /// <summary>
        /// 选择本地文件存放路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            string prePath = label2.Text;
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.label2.Text = path.SelectedPath;
            loggerConfig.Trace($"【变更存放路径】\n原先存放路径：{prePath}\n路径已变更为：{this.label2.Text}");
        }

        private void enterButton_Click(object sender, EventArgs e)
        {
            directoryTreeView.Nodes.Clear();// 每次确定时需要刷新内容
            string inputText = label2.Text; // 获得输入框的内容
            // 文件路径存在
            if (Directory.Exists(inputText))
            {
                TreeNode rootNode = new TreeNode(inputText); // 创建树节点
                directoryTreeView.Nodes.Add(rootNode); // 加入视图
                FindDirectory(inputText, rootNode);  //通过递归函数进行目录的遍历
            }

        }

        // 递归函数 遍历当前目录
        void FindDirectory(string nowDirectory, TreeNode parentNode)
        {
            try  // 当文件目录不可访问时，需要捕获异常
            {
                // 获取当前目录下的所有文件夹数组
                string[] directoryArray = Directory.GetFiles(nowDirectory);
                if (directoryArray.Length > 0)
                {
                    foreach (string item in directoryArray)
                    {
                        // 遍历数组，将节点添加到父亲节点的
                        string str = Path.GetFileNameWithoutExtension(item);
                        TreeNode node = new TreeNode(str);
                        parentNode.Nodes.Add(node);
                        //FindDirectory(item, node);
                    }
                }
            }
            catch (Exception)
            {
                parentNode.Nodes.Add("禁止访问");
            }
        }

        private void directoryTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (directoryTreeView.SelectedNode.FullPath == directoryTreeView.SelectedNode.Text)
            {
                string table = directoryTreeView.SelectedNode.Text;//数据表名

                directoryTreeView.SelectedNode.Expand();//展开选中的节点
            }
            else
            {
                int count = directoryTreeView.SelectedNode.Text.Length;             //获取选中节点字符长度
                string str = directoryTreeView.SelectedNode.FullPath;               //获取选中节点从父节点到目标节点的路径

                textBoxPath.Text = str;
            }
        }

        /// <summary>
        /// 变更工单
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");
            if (table1.Rows.Count > 0)
            {
                string sql = "update [SytemSet] set [wordNo]='" + textBox6.Text + "' where [ID] = '1'";
                var result = mdb.Change(sql);
                if (result == true)
                {
                    MessageBox.Show("变更成功");
                }
            }
            mdb.CloseConnection();
        }

        #region --------------- 用户管理 ---------------

        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            // OP = 1, PE = 2, ADM = 3, DEV = 4, QE = 5, ME = 6
            int[] OpAuth = { 1, 2, 5, 6 };                      // 只能注册操作员权限
            string[] restrictedUsers = { "PE", "QE", "ME" };    // 只能由 ADM 进行管理的权限
            int[] topLevel = { 3, 4 };                          // 最高权限用户
            // 获取数据库中，用户工号的集合
            string[] userID = userCollection.AsEnumerable()
                                  .Select(row => row.Field<string>("工号"))
                                  .ToArray();

            // 确认当前用户是否为 OP、PE、QE、ME
            bool isOpAuth = OpAuth.Contains(Access);
            // 确认当前用户是否为 ADM、DEV
            bool isTopLevel = topLevel.Contains(Access);

            // 检查必填项是否为空
            if (UID.Text.Trim().Length < 1 || textBox22.Text.Trim().Length < 1 || UPWD.Text.Trim().Length < 1)
            {
                MessageBox.Show("工号、姓名、密码为必填项！", "提示");
                return;
            }

            // 确保 OP、PE、QE、ME 用户只能注册 OP 权限
            if (isOpAuth && UTYPE.Text != "OP")
            {
                MessageBox.Show("您只能注册操作员权限，请重新选择！", "提示");
                return;
            }

            // 确保 PE、QE、ME 用户只能由 ADM、DEV 等级用户进行管理
            if (restrictedUsers.Contains(UTYPE.Text) && !isTopLevel)
            {
                MessageBox.Show("您没有权限更改此用户等级，请重新选择！", "提示");
                return;
            }

            // 确保 OP 可以被注册
            if (UTYPE.Text == "OP" && !userID.Contains(UID.Text))
            {

            }
            // 防止除最高权限以外的权限将其他权限被篡改为 OP 权限
            else if (UTYPE.Text == "OP" && lblCurrentSelected.Text != "OP" && !isTopLevel)
            {
                MessageBox.Show("您没有权限更改此用户等级，请重新选择！", "提示");
                return;
            }

            string sql = "";
            mdb = new mdbDatas(userFileuRL);
            DataTable table1 = mdb.Find("select * from Users where 工号 = '" + UID.Text + "'");

            if (table1.Rows.Count > 0)
            {
                //if (comboBox3.SelectedIndex == 1)
                //{
                if (tbxBrandID.Text.Length > 0)
                {
                    DataTable table = mdb.Find("select * from Users where [厂牌UID] = '" + tbxBrandID.Text + "' and [厂牌UID] <> '' and [工号] <> '" + UID.Text + "' ");
                    if (table.Rows.Count > 0)
                    {
                        if (tbxBrandID.Text == table.Rows[0]["厂牌UID"].ToString())
                        {
                            MessageBox.Show("不能重复注册、请检查厂牌UID是否重复！");
                            return;
                        }

                    }
                }
                sql = "update [Users] set [用户密码]='" + UPWD.Text + "',[用户权限]='" + UTYPE.Text + "',[用户名]='" + textBox22.Text + "',[厂牌UID]='" + tbxBrandID.Text + "'" +
                             " where [工号] = '" + UID.Text + "'";
                //}
                //else 
                //{
                //    sql = "update [Users] set [工号]='" + UID.Text + "',[用户密码]='" + UPWD.Text + "',[用户权限]='" + UTYPE.Text + "',[登录方式]='" + GetAuthorityByCode(comboBox3.SelectedIndex) + "',[用户名]='" + textBox22.Text + "',[厂牌UID]=' '" +
                //              " where [工号] = '" + UID.Text + "'";
                //}

                var result = mdb.Change(sql);
                if (result == true)
                {
                    MessageBox.Show("修改成功");
                }
            }
            else
            {



                //if (comboBox3.SelectedIndex == 1)
                //{
                if (tbxBrandID.Text.Length > 0)
                {
                    DataTable table = mdb.Find("select * from Users where [厂牌UID] = '" + tbxBrandID.Text + "' and [厂牌UID] <> '' ");
                    if (table.Rows.Count > 0)
                    {
                        MessageBox.Show("不能重复注册、请检查厂牌UID是否重复！");
                        return;
                    }
                }

                sql = "insert into Users ([工号],[用户密码],[用户权限],[用户名],[厂牌UID]) values ('" + UID.Text + "','" + UPWD.Text + "','" + UTYPE.Text + "','" + textBox22.Text + "','" + tbxBrandID.Text + "')";
                //}
                //else 
                //{
                //     sql = "insert into Users ([工号],[用户密码],[用户权限],[登录方式],[用户名]) values ('" + UID.Text + "','" + UPWD.Text + "','" + UTYPE.Text + "','" + GetAuthorityByCode(comboBox3.SelectedIndex) + "','" + textBox22.Text + "')";
                //}


                bool result = mdb.Add(sql.ToString());
                if (result == true)
                {
                    MessageBox.Show("新增成功");
                }

            }
            button6_Click(null, null);
            mdb.CloseConnection();

            string modifyInfo = SqlToJsonConverter.ConvertToJSON(sql);
            //string toLog = SqlToJsonConverter.ConvertToLog(sql);
            string final = SqlToJsonConverter.ConvertToCustomLog(modifyInfo);
            loggerAccount.Trace(final);
        }

        /// <summary>
        /// 删除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            // OP = 1, PE = 2, ADM = 3, DEV = 4, QE = 5, ME = 6
            int currentUserAccess = Access;
            int delUserAccess = GetAccessLevelFromUtype(UTYPE.Text);

            if (new[] { 2, 5, 6 }.Contains(delUserAccess))  // PE, QE, ME
            {
                if (currentUserAccess != 3 && currentUserAccess != 4)
                {
                    MessageBox.Show("当前权限仅有管理员才能删除！", "提示");
                    return;
                }
            }

            if (delUserAccess == 3)   // ADM
            {
                if (currentUserAccess != 3 && currentUserAccess != 4)
                {
                    MessageBox.Show("当前权限仅有管理员才能删除！", "提示");
                    return;
                }
            }

            //获取当前的行
            int rowindex = dataGridView1.CurrentRow.Index;

            //将获取到的当前行转为id
            string id = (string)dataGridView1.Rows[rowindex].Cells[4].Value;

            // 连接到数据库
            mdb = new mdbDatas(userFileuRL);

            // 删除前先执行查找
            string selectSql = $"SELECT [用户名], [用户权限] FROM [Users] WHERE [工号] = '{id}'";
            DataTable userInfo = mdb.Find(selectSql);

            string username = "";
            string permissions = "";
            if (userInfo.Rows.Count > 0)
            {
                username = userInfo.Rows[0]["用户名"].ToString();
                permissions = userInfo.Rows[0]["用户权限"].ToString();
            }

            string sql = $" DELETE FROM [Users] WHERE [工号] = '{id}' ";
            bool bl = mdb.Del(sql);
            if (bl == true)
            {
                MessageBox.Show("删除成功");
                // 记录日志
                string delInfo = $"【用户删除】\n工号：{id} | 姓名：{username} | 权限：{permissions}";
                loggerAccount.Trace(delInfo);
            }
            button6_Click(null, null);
            mdb.CloseConnection();
        }

        DataTable userCollection;
        /// <summary>
        /// 刷新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(userFileuRL);
            userCollection = mdb.Find("select 用户名,用户密码,用户权限,厂牌UID,工号 from Users where 用户权限 <> 'DEV'");
            userInfoEntities = DataConverter.ConvertDataTableToList(userCollection);
            dataGridView1.DataSource = userCollection;
            mdb.CloseConnection();
        }

        /// <summary>
        /// 登录方式转换
        /// </summary>
        /// <param name="authorityName"></param>
        /// <returns></returns>
        public string GetAuthorityByCode(int authorityName)
        {
            string aCode = "";
            switch (authorityName)
            {
                case 0:
                    aCode = "密码";
                    break;
                case 1:
                    aCode = "刷卡";
                    break;
                default:
                    aCode = "密码";
                    break;
            }
            return aCode;
        }

        /// <summary>
        /// 编码转换名称
        /// </summary>
        /// <param name="authorityName"></param>
        /// <returns></returns>
        public int GetCodeByAuthority(string Code)
        {
            int aCode = 0;
            switch (Code)
            {
                case "密码":
                    aCode = 0;
                    break;
                case "刷卡":
                    aCode = 1;
                    break;
                default:
                    aCode = 0;
                    break;
            }
            return aCode;
        }

        /// <summary>
        /// 修改登录次数和登录时间
        /// </summary>
        public void UpLoginInfo()
        {
            mdb = new mdbDatas(userFileuRL);
            DataTable table = mdb.Find("select 登录次数 from Users where [工号] = '" + LoginUser + "'");
            if (table.Rows.Count > 0)
            {
                int i = 0;
                if (!string.IsNullOrWhiteSpace(table.Rows[0][0].ToString()))
                {
                    i = int.Parse(table.Rows[0][0].ToString());
                }
                string sql = "update [Users] set [最后登录时间]='" + DateTime.Now.ToString() + "',[登录次数]='" + (i + 1) + "' where [工号] = '" + LoginUser + "'";

                var result = mdb.Change(sql);
                //if (result == true)
                //{
                //    MessageBox.Show("修改成功");
                //}
            }
            mdb.CloseConnection();
        }

        /// <summary>
        /// 获取用户当前选中的权限类型
        /// </summary>
        /// <param name="utype"></param>
        /// <returns></returns>
        private int GetAccessLevelFromUtype(string utype)
        {
            switch (utype)
            {
                case "OP": return 1;
                case "PE": return 2;
                case "ADM": return 3;
                case "DEV": return 4;
                case "QE": return 5;
                case "ME": return 6;
                default: return 0;
            }
        }

        /// <summary>
        /// 选中行事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                this.UID.Text = this.dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
                this.UPWD.Text = this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                this.textBox22.Text = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                this.tbxBrandID.Text = this.dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
                this.UTYPE.Text = this.dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                this.lblCurrentSelected.Text = UTYPE.Text;
                //this.comboBox3.SelectedIndex = GetCodeByAuthority(this.dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString());


            }
        }

        #endregion

        #region --------------- 设置验证条码与读取PLC ---------------

        /// <summary>
        /// 设置验证条码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button22_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");
            if (table1.Rows.Count > 0)
            {
                string comtb = "0";
                if (!string.IsNullOrWhiteSpace(comboBox2.Text))
                {
                    comtb = comboBox2.SelectedValue.ToString();
                }
                string sql = $"update [SytemSet] set [faults]='{comtb}', Workstname ='{checkBox6.Checked}' where [ID] = '1'";
                var result = mdb.Change(sql);
                if (result == true)
                {
                    label50.Text = comtb;
                    MessageBox.Show("保存成功");

                    // 记录操作日志
                    bool isChecked = checkBox6.Checked;
                    string readPlcStatus = isChecked ? "是" : "否";
                    string msgSave = $"【配方操作保存成功】\n是否读取PLC：{readPlcStatus}\n当前配方号：{label50.Text}";
                    loggerConfig.Trace(msgSave);
                }
            }

            mdb.CloseConnection();
        }
        /// <summary>
        /// 发给PLC
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button25_Click_1(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                MessageBox.Show("读取PLC中，不能发送！");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(comboBox2.Text))
                {
                    MessageBox.Show("请选择条码验证信息！");
                    return;
                }
                try
                {
                    KeyenceMcNet.Write("D1204", 1);
                    KeyenceMcNet.Write("D1206", int.Parse(comboBox2.SelectedValue.ToString()));
                    MessageBox.Show("发送PLC成功！");
                }
                catch (Exception ex)
                {
                    LogMsg(ex.ToString());
                    MessageBox.Show("发送PLC失败！");
                }
            }
        }
        #endregion

        #region --------------- 分页 ---------------
        Pager pager = null;
        //将分页面数据显示出来
        private void PageLoad()
        {
            dataGridViewDynamic1.DataSource = pager.LoadPage();//显示数据
            this.textBox13.Text = pager.pageSize.ToString();//每页的条
            this.label34.Text = pager.currentPage + "/" + pager.pageCount;//当前页0/0
            this.label36.Text = pager.recordCount.ToString();//  共条
        }

        //首页
        private void button4_Click(object sender, EventArgs e)
        {
            if (pager != null)
            {
                pager.currentPage = 1;
                PageLoad();//显示分页数据
            }
        }
        //上一页
        private void button9_Click_1(object sender, EventArgs e)
        {
            if (pager != null)
            {
                pager.currentPage--;
                PageLoad();//显示分页数据
            }
        }
        //下一页
        private void button11_Click_1(object sender, EventArgs e)
        {
            if (pager != null)
            {
                pager.currentPage++;
                PageLoad();//显示分页数据
            }
        }
        //尾页
        private void button12_Click(object sender, EventArgs e)
        {
            if (pager != null)
            {
                pager.currentPage = pager.pageCount;
                PageLoad();//显示分页数据
            }
        }
        //跳转
        private void button14_Click(object sender, EventArgs e)
        {
            // pager.currentPage = this.textBox16.Text;
            if (pager != null)
            {
                int i;
                if (int.TryParse(this.textBox16.Text, out i))
                {
                    pager.currentPage = i;
                    PageLoad();//显示分页数据
                }
            }

        }
        //重新分页
        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            if (pager != null)
            {
                int i;
                if (int.TryParse(this.textBox13.Text, out i))
                {
                    if (i != 0) pager.pageSize = i;
                    pager.fenye();//分页
                    PageLoad();//显示分页数据
                }

            }
        }



        #endregion

        #region --------------- 看板数据库设置 ---------------

        Socket socketSend;
        private bool IsStart = false;
        private Action<string> ShowMsgAction;
        DataTable vubPartsTable;                // 易损件表
        DataTable faultsTable;                  // 故障信息表
        //string[] boardName = new string[] { };//易损件
        //string[] boardCode = new string[] { };
        //string[] boardTeory = new string[] { };//理论值
        //string[] boardPosition = new string[] { };//位置

        /// <summary>
        /// 加载看板界面相关参数
        /// </summary>
        private void SYS_Socket_Mo()
        {
            button18_Click(null, null);

            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSocket where ID = 1");

            // 加载看板界面参数
            for (int i = 0; i < table1.Rows.Count; i++)
            {
                for (int j = 0; j < table1.Columns.Count; j++)
                {
                    tbxBulletinIP.Text = table1.Rows[i]["IP"].ToString();           // 看板IP
                    tbxBulletinPort.Text = table1.Rows[i]["Port"].ToString();       // 看板端口号

                    string device = table1.Rows[i]["Devicestatu"].ToString();
                    if (device == "True") cbxOpenBulletin.Checked = true;           // 看板状态

                    string cokPosi = table1.Rows[i]["BoardPosition"].ToString();
                    if (cokPosi == "True") checkBox2.Checked = true;                // 框架螺栓组件

                    tbxStationName.Text = table1.Rows[i]["BoardTheory"].ToString(); // 工位名称
                    textBox30.Text = table1.Rows[i]["BoardName"].ToString();        // 工位名称集合

                    textBox39.Text = table1.Rows[i]["FaultCode"].ToString();        // 故障起始点位
                    textBox40.Text = table1.Rows[i]["FaultLeng"].ToString();        // 长度
                }
            }

            mdb.CloseConnection();

            // 创建一个新的列对象并设置其属性保存
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
            buttonColumn.HeaderText = "操作";
            buttonColumn.Text = "保存";
            buttonColumn.Name = "btnCol";
            buttonColumn.DefaultCellStyle.NullValue = "保存";
            dataGridView2.Columns.Add(buttonColumn);

            // 再次创建一个新的列对象并设置其属删除
            DataGridViewButtonColumn anotherButtonColumn = new DataGridViewButtonColumn();
            anotherButtonColumn.HeaderText = "操作"; // 第二个按钮的标题文本
            anotherButtonColumn.Name = "btnCol2"; // 第二个按钮的名称
            anotherButtonColumn.DefaultCellStyle.NullValue = "删除";
            dataGridView2.Columns.Add(anotherButtonColumn);

            DataGridViewButtonColumn butnCo = new DataGridViewButtonColumn();
            butnCo.HeaderText = "操作";
            butnCo.Text = "保存";
            butnCo.Name = "btnCol";
            butnCo.DefaultCellStyle.NullValue = "保存";
            dataGridView3.Columns.Add(butnCo);

            // 再次创建一个新的列对象并设置其属删除
            DataGridViewButtonColumn anotrButCo = new DataGridViewButtonColumn();
            anotrButCo.HeaderText = "操作"; // 第二个按钮的标题文本
            anotrButCo.Name = "btnCol2"; // 第二个按钮的名称
            anotrButCo.DefaultCellStyle.NullValue = "删除";
            dataGridView3.Columns.Add(anotrButCo);
        }

        /// <summary>
        /// 刷新易损件表和故障信息表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button18_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);

            // 查询易损件表
            vubPartsTable = mdb.Find($"select [ID] as 编号, [WorkID] as 工位ID, [BoardPosition] as 易损件所在的位置, " +
                $"[BoardName] as 易损件的名称, [BoardTheory] as 理论使用次数PLC点位, " +
                $"[BoardCode] as 已经使用的PLC点位 from VulnbleParts");

            vubPartsTable.Columns["编号"].AutoIncrement = true;
            vubPartsTable.Columns["编号"].ReadOnly = true;
            vubPartsTable.Columns["编号"].AutoIncrementSeed = 0;

            int maxPid = 0;
            if (vubPartsTable.Rows.Count > 0)
            {
                maxPid = vubPartsTable.AsEnumerable().Max(row => row.Field<int>("编号"));
            }

            DataRow newRow = vubPartsTable.NewRow();
            newRow["编号"] = maxPid;

            // 刷新表格
            dataGridView2.DataSource = vubPartsTable;

            // 查询故障信息表
            faultsTable = mdb.Find("select ID as 编号, [WorkID] as 工位ID, [CodeID] as 故障点位, Faults as 故障描述 from SytemFaults");
            // 刷新表格
            dataGridView3.DataSource = faultsTable;

            /* if (textBox19.Text.Length > 0)
             {
                 boardName = textBox20.Text.ToString().Split(new char[] { '|' });//易损件名称
                 boardCode = textBox19.Text.ToString().Split(new char[] { '|' });//采取PLC易损件使用的点
                 boardTeory = textBox21.Text.ToString().Split(new char[] { '|' });//易损件的理论值
                 boardPosition = textBox25.Text.ToString().Split(new char[] { '|' });//易损件所在该机台的位置
             }*/
            mdb.CloseConnection();
        }

        /// <summary>
        /// 易损件 1.删除2.保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            { return; }

            // 需要记录到日志的字段
            string[] logFields = { "WorkID", "BoardPosition", "BoardName", "BoardTheory", "BoardCode" };
            // 定义字段与别名的映射
            Dictionary<string, string> fieldAliases = new Dictionary<string, string>
            {
                { "WorkID", "工位序号" },
                { "BoardPosition", "易损件所在的位置" },
                { "BoardName", "易损件的名称" },
                { "BoardTheory", "理论使用次数PLC点位" },
                { "BoardCode", "已经使用的PLC点位" },
            };

            // 删除
            DeleteRowFromDataGridView<int>(dataGridView2, e, "VulnbleParts", "ID", logFields, 2, path4, "btnCol2", fieldAliases, "易损件数据");

            // 保存
            if (dataGridView2.Columns[e.ColumnIndex].Name == "btnCol")
            {
                //说明点击的列是DataGridViewButtonColumn列
                DataGridViewColumn column = dataGridView2.Columns[e.ColumnIndex];
                string pid = this.dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                string workid = this.dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
                string bdPosition = this.dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                string bdName = this.dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString();
                string bdTheory = this.dataGridView2.Rows[e.RowIndex].Cells[6].Value.ToString();
                string bdCode = this.dataGridView2.Rows[e.RowIndex].Cells[7].Value.ToString();

                mdb = new mdbDatas(path4);
                DataTable table1 = mdb.Find($"select * from VulnbleParts where [ID] = {pid}");

                // 修改
                if (table1.Rows.Count > 0)
                {
                    DataRow row = table1.Rows[0];

                    // 修改前的详细数据
                    string logDetail = $"编号：{row["ID"]} | 工位序号：{row["WorkID"]} | " +
                        $"易损件所在的位置：{row["BoardPosition"]} | 易损件的名称：{row["BoardName"]} | " +
                        $"理论使用次数PLC点位：{row["BoardTheory"]} | 已经使用的PLC点位：{row["BoardCode"]}";

                    string sql = $"update [VulnbleParts] set [WorkID] = '{workid}', [BoardPosition] = '{bdPosition}', " +
                    $"[BoardName] = '{bdName}', [BoardTheory] = '{bdTheory}', [BoardCode] = '{bdCode}' where [ID] = {pid}";

                    var result = mdb.Change(sql);
                    if (result == true)
                    {
                        MessageBox.Show("修改成功");
                        string modifyInfo = $"【易损件数据修改成功】\n修改前的详细信息：\n{logDetail}\n修改后的详细信息：\n" +
                            $"编号：{pid} | 工位序号：{workid} | 易损件所在的位置：{bdPosition} | 易损件的名称：{bdName} | " +
                            $"理论使用次数PLC点位：{bdTheory} | 已经使用的PLC点位：{bdCode}";
                        loggerConfig.Trace(modifyInfo);
                    }
                }
                // 新增
                else
                {
                    string sql = $"insert into VulnbleParts ([ID], [WorkID], [BoardPosition], [BoardName], [BoardTheory], [BoardCode]) " +
                                 $"values ({pid}, '{workid}', '{bdPosition}', '{bdName}', '{bdTheory}', '{bdCode}')";
                    bool result = mdb.Add(sql.ToString());
                    if (result == true)
                    {
                        MessageBox.Show("新增成功");
                        string insertInfo = $"【易损件数据新增成功】\n新增详情：\n" +
                            $"编号：{pid} | 工位序号：{workid} | 易损件所在的位置：{bdPosition} | 易损件的名称：{bdName} | " +
                            $"理论使用次数PLC点位：{bdTheory} | 已经使用的PLC点位：{bdCode}";
                        loggerConfig.Trace(insertInfo);
                    }
                }

                mdb.CloseConnection();
            }

        }

        /// <summary>
        /// 故障信息 1.删除2.保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            { return; }

            // 需要记录日志的字段
            string[] logField = { "WorkID", "CodeID", "Faults" };
            // 定义字段与别名的映射
            Dictionary<string, string> fieldAliases = new Dictionary<string, string>
            {
                { "WorkID", "工位序号" },
                { "CodeID", "故障点位" },
                { "Faults", "故障描述" },
            };

            // 删除
            DeleteRowFromDataGridView<string>(dataGridView3, e, "SytemFaults", "ID", logField, 2, path4, "btnCol2", fieldAliases, "故障信息");

            // 保存
            if (dataGridView3.Columns[e.ColumnIndex].Name == "btnCol")
            {
                DataGridViewColumn column = dataGridView3.Columns[e.ColumnIndex];
                string pid = this.dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString();
                string workID = this.dataGridView3.Rows[e.RowIndex].Cells[3].Value.ToString();
                string codeID = this.dataGridView3.Rows[e.RowIndex].Cells[4].Value.ToString();
                string funmae = this.dataGridView3.Rows[e.RowIndex].Cells[5].Value.ToString();

                if (string.IsNullOrWhiteSpace(pid))
                {
                    MessageBox.Show("故障信息编号不能为空！");
                    return;
                }

                mdb = new mdbDatas(path4);
                DataTable table1 = mdb.Find($" select * from SytemFaults where [ID] = '{pid}' ");

                // 修改
                if (table1.Rows.Count > 0)
                {
                    DataRow row = table1.Rows[0];

                    // 修改前的详细数据
                    string logDetail = $"修改前的详细信息：\n" +
                        $"编号：{row["ID"]} | 工位序号：{row["WorkID"]} | " +
                        $"故障点位：{row["CodeID"]} | 故障描述：{row["Faults"]}";

                    string sql = $"update [SytemFaults] set [WorkID] = '{workID}', [CodeID] = '{codeID}', " +
                        $"[Faults] = '{funmae}' where [ID] = '{pid}'";

                    var result = mdb.Change(sql);
                    if (result == true)
                    {
                        MessageBox.Show("修改成功");
                        string updateInfo = $"修改后的详细信息：\n" +
                            $"编号：{pid} | 工位序号：{workID} | " +
                            $"故障点位：{codeID} | 故障描述：{funmae}";
                        loggerConfig.Trace($"【故障信息修改成功】\n{logDetail}\n{updateInfo}");
                    }
                }

                // 新增
                else
                {
                    string sql = $"insert into SytemFaults ([ID],[WorkID],[CodeID],[Faults]) " +
                        $"values ('{pid}', '{workID}', '{codeID}', '{funmae}')";

                    bool result = mdb.Add(sql.ToString());
                    if (result == true)
                    {
                        MessageBox.Show("新增成功");
                        string insertInfo = $"【故障信息新增成功】\n新增详情：\n" +
                            $"编号：{pid} | 工位序号：{workID} | " +
                            $"故障点位：{codeID} | 故障描述：{funmae}";
                        loggerConfig.Trace(insertInfo);
                    }
                }

                mdb.CloseConnection();
            }

        }

        /// <summary>
        /// 看板设置界面中的 “保存” 按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button17_Click_1(object sender, EventArgs e)
        {
            if (tbxBulletinIP.Text == String.Empty || tbxBulletinPort.Text == String.Empty)
            {
                MessageBox.Show("当前界面内容均为必填项、请先填写完善");
                return;
            }

            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from SytemSocket where ID = 1");

            if (table1.Rows.Count > 0)
            {
                string sql = $"update [SytemSocket] set [IP] = '{tbxBulletinIP.Text}', [Port] = '{tbxBulletinPort.Text}', " +
                    $"[Devicestatu] = '{cbxOpenBulletin.Checked}', [BoardName] = '{textBox30.Text}', " +
                    $"[BoardTheory] = '{tbxStationName.Text}', [BoardPosition] = '{checkBox2.Checked}', " +
                    $"[FaultCode] = '{textBox39.Text}', [FaultLeng] = '{textBox40.Text}' where [ID] = 1";

                var result = mdb.Change(sql);
                if (result == true)
                {
                    MessageBox.Show("保存成功");
                }
            }
            mdb.CloseConnection();
        }

        #endregion

        #region --------------- 看板Socket连接服务器 ---------------

        private System.Net.Sockets.Socket socket;

        /// <summary>
        /// 连接
        /// </summary>
        /// <param name="ipAddress"></param>
        /// <param name="port"></param>
        public async void Connect()
        {
            Invoke(new Action(() =>
            {
                lblDashboardStatus.ForeColor = R;
                //  label43.Text = "看板状态：未连接";
            }));

            await Task.Run(() =>
                {
                    while (true)
                    {
                        try
                        {
                            // Internet 协议、字节流、IPv4连接
                            socket = new Socket(
                            AddressFamily.InterNetwork,
                            SocketType.Stream,
                            ProtocolType.IP);
                            IPAddress ip = IPAddress.Parse(tbxBulletinIP.Text);
                            IPEndPoint point = new IPEndPoint(ip, Convert.ToInt32(tbxBulletinPort.Text));
                            socket.Connect(point);
                            IsStart = true;

                            if (IsStart)
                            {
                                //读取消息
                                System.Threading.Thread thread = new System.Threading.Thread(Received);
                                thread.IsBackground = true;
                                thread.Start();

                                //发送心跳包
                                System.Threading.Thread thread1 = new System.Threading.Thread(Sendheartbeat);
                                thread1.IsBackground = true;
                                thread1.Start();

                                //读取工单号等生产数据发送
                                System.Threading.Thread thread2 = new System.Threading.Thread(SedMegin);
                                thread2.IsBackground = true;
                                thread2.Start();



                                Invoke(new Action(() =>
                                {
                                    lblDashboardStatus.ForeColor = G;
                                    //  label43.Text = "看板状态：已连接";
                                }));

                                break;
                            }


                        }
                        catch (Exception ex)
                        {
                            Thread.Sleep(50000);
                            ShowMsg("Reconnection failed: " + ex.Message);
                        }

                    }
                });
        }

        /// <summary>
        /// 客户端Socket连接服务器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button16_Click(object sender, EventArgs e)
        {

            if (cbxOpenBulletin.Checked == false)
            {

                IsStart = false;
                lblDashboardStatus.ForeColor = B;
                // label43.Text = "看板状态：未启用";
                return;
            }
            Connect();

            //await Task.Run(() => { 
            //try
            //{
            //    //创建客户端Socket，获得远程ip和端口号
            //    socketSend = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //    IPAddress ip = IPAddress.Parse(textBox18.Text);
            //    IPEndPoint point = new IPEndPoint(ip, Convert.ToInt32(textBox12.Text));
            //    socketSend.Connect(point);
            //    ShowMsg("连接成功!");
            //    //开启新的线程，不停的接收服务器发来的消息
            //    Thread c_thread = new Thread(Received);
            //    c_thread.IsBackground = true;
            //    c_thread.Start();
            //    //每次连接都将工位号发给服务端
            //    Task thread = new Task(SedMegin);//读取工单号等生产数据
            //    thread.Start();
            //    // Send("0+"+textBox24.Text);
            //    IsStart = true;

            //    // ShowBtnState();
            //}
            //catch (Exception ex)
            //{
            //    IsStart = false;

            //}
            //});

        }

        /// <summary>
        /// 接收线程
        /// </summary>
        private void Received()
        {
            while (IsStart)
            {
                try
                {
                    //if (!socket.Connected) return;
                    byte[] buffer = new byte[1024 * 1024 * 3];
                    //实际接收到的有效字节数
                    int len = socket.Receive(buffer);
                    if (len == 0)
                    {
                        break;
                    }
                    string str = Encoding.UTF8.GetString(buffer, 0, len);
                    ReceiveData(str);
                    this.BeginInvoke(ShowMsgAction, socket.RemoteEndPoint + ":" + str);
                    ShowMsg(socket.RemoteEndPoint + ":" + str);
                }
                catch (Exception ex)
                {
                    IsStart = false;
                    Console.WriteLine(ex.Message);
                    socket.Close();
                    Connect();
                }
            }
        }

        /// <summary>
        /// 定时发送心跳包
        /// </summary>
        private async void Sendheartbeat()
        {
            await Task.Run(() =>
            {
                // 定时发送心跳包并检测连接状态
                while (IsStart)
                {
                    try
                    {
                        // 发送心跳包
                        byte[] heartbeatData = System.Text.Encoding.UTF8.GetBytes("heartbeat");
                        socket.Send(heartbeatData);
                        Thread.Sleep(100000);

                    }
                    catch (Exception ex)
                    {
                        IsStart = false;
                        Console.WriteLine(ex.Message);
                        socket.Close();

                        Connect();
                        //while (true)
                        //{
                        //    try
                        //    {
                        //        if (!IsStart)
                        //        {
                        //            Connect();
                        //            if (IsStart)
                        //            {
                        //                break;
                        //            }

                        //        }
                        //        Thread.Sleep(5000);
                        //    }
                        //    catch (Exception ex2)
                        //    {
                        //        ShowMsg("Connect failed: " + ex2.Message);
                        //    }
                        //}

                    }
                }
            });
        }

        public void SedMegin()
        {
            Invoke(new Action(() =>
            {

                Thread.Sleep(50);
                Application.DoEvents();
                Send("0+" + tbxStationName.Text);
                Thread.Sleep(50);
                Application.DoEvents();

                Thread.Sleep(200);
                Application.DoEvents();
                Send("5+" + textBox3.Text);
            }));
        }

        public void SendReceived()
        {
            Invoke(new Action(() =>
            {
                string bulletindata = "";
                //统计

                //3+机台名称+工单号+工单数量+完成数量+
                //完成率+合格率+整体节拍+生产产品数量（总数）
                //+ 工序时间+利用时间+负荷时间+直通率+成品名称
                string shotji = "3+" + textBox3.Text + "+" + textBox6.Text + "+" + D1084 + "+" + D1086 + "+"
                    + complte + "+" + Paate + "+" + D1090 + "+" + D1080
                        + "+" + prodproces + "+" + ustim + "+" + loadti
                        + "+" + shootthgh + "+" + tbxProductName.Text;
                bulletindata += shotji;
                //ShowMsg(shotji);
                //Send(shotji);
                //读取易损件使用数目vulpalist
                if (vubPartsTable != null)
                {
                    foreach (DataRow row in vubPartsTable.Rows)
                    {
                        string vulpa = "  ";
                        string kordnme = " ";
                        string vuboardTheory = row["理论使用次数PLC点位"].ToString();
                        string vuboard = row["已经使用的PLC点位"].ToString();
                        var boardTheory = KeyenceMcNet.ReadInt32(vuboardTheory).Content;
                        OperateResult<int> readss = KeyenceMcNet.ReadInt32(vuboard);
                        if (readss.IsSuccess)
                        {
                            vulpa = readss.Content.ToString();
                            kordnme = CodeNum.WorkIDNm(row["工位ID"].ToString(), stationName);
                            //输出：
                            //4+易损件所在工位+机台名称+ 易损件所在位置+
                            //易损件名称+易损件理论使用次数易损件已使用次数
                            string shotyisjian = "4+" + tbxStationName.Text + "+" + label54.Text + "+" + row["易损件所在的位置"] + "+"
                               + row["易损件的名称"] + "+" + boardTheory + "+" + vulpa;
                            bulletindata += "|" + shotyisjian;

                            //ShowMsg(shotyisjian);
                            //Send(shotyisjian);
                        }
                    }
                }
                //2+工位名称+当前工单号+产品条码+
                //操作人员+则式时间+
                //则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值


                string chesAAA = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                                       + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                                           "+" + Value[9999] + "+" + D1090 + "+产品名称+" + "  " + "+" + "  " + "+" + tbxProductName.Text;
                bulletindata += "|" + chesAAA;
                //ShowMsg(chesAAA);
                //Send(chesAAA);
                string[] frorckstr = textBox41.Text.Split('+');
                for (int i = 0; i < frorckstr.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(frorckstr[i]))
                    {
                        string chesBBB = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                                        + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                                            "+" + Value[9999] + "+" + D1090 + "+工装编号" + (i + 1) + "+" + "  " + "+" + "  " + "+" + frorckstr[i];
                        bulletindata += "|" + chesBBB;
                    }
                    // ShowMsg(chesBBB);
                    // Send(chesBBB);
                }
                string[] prodrow = CodeNum.CodeMafror(comboBox2.Text, codesDataM);
                for (int i = 0; i < prodrow.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(prodrow[i]))
                    {
                        string chesCCC = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                                           + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                                               "+" + Value[9999] + "+" + D1090 + "+产品物料号" + (i + 1) + "+" + " " + "+" + "  " + "+" + prodrow[i];
                        bulletindata += "|" + chesCCC;
                    }
                    // ShowMsg(chesCCC);
                    //Send(chesCCC);
                }
                if (
               list.Count > 0 && AstrName.Count > 0)
                {

                    for (int i = 0; i < list.Count; i++)
                    {
                        if (!list[i].Equals("null"))
                        {

                            //输出：
                            //2+工位名称+当前工单号+产品条码+
                            //操作人员+则式时间+
                            //则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值
                            if (maxValue[i].Equals("NO") && minValue[i].Equals("NO") && testResult[i].Equals("NO"))
                            {
                                string cheshixm = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                                        + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                                            "+" + Value[9999] + "+" + D1090 + "+" + AstrName[i] + "+" + "" + "+" + "" + "+" + list[i];
                                bulletindata += "|" + cheshixm;
                                // ShowMsg(cheshixm);
                                // Send(cheshixm);
                            }
                            else
                            {
                                string cheshixm1 = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                                    + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                                        "+" + Value[9999] + "+" + D1090 + "+" + AstrName[i] + "+" + maxlist[i] + "+" + minlist[i] + "+" + list[i];
                                bulletindata += "|" + cheshixm1;
                                //ShowMsg(cheshixm1);
                                //Send(cheshixm1);
                            }

                        }
                    }
                }
                else
                {
                    //输出：
                    //2+工位名称+当前工单号+产品条码+
                    //操作人员+则式时间+
                    //则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值
                    //注意：(则试项名称+ 则试项上限+则试项下限+测式项实际值)为空


                    string cheshixm2 = "2+" + tbxStationName.Text + "+" + textBox6.Text + "+" + barcodeData + "+"
                     + LoginUser.ToString() + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                            "+" + Value[9999] + "+" + D1090 + "+" + "测试总结果" + "+" + "   " + "+" + "    " + "+" + Value[9999];
                    bulletindata += "|" + cheshixm2;

                    //ShowMsg(cheshixm2);
                    // Send(cheshixm2);

                }
                ShowMsg(bulletindata);
                Send(bulletindata);
            }));
        }

        public static readonly object objSync = new object();

        /// <summary>
        /// 发信息
        /// 1+故障所在工位+机台名称+ 故障状态+故障的描述+触发故障的开始时间 
        /// 1+故障所在工位+机台名称+ 故障状态+故障的描述+触发故障的结束时间
        /// 2+工位名称+当前工单号+产品条码+操作人员+则式时间+则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值
        /// 3+工位+工单号+工单数量+完成数量+完成率+合格率+整体节拍+生产产品数量（总数）+ 工序时间+利用时间+负荷时间
        /// 4+易损件所在工位+机台名称+ 易损件所在位置+易损件名称+易损件理论使用次数+易损件已使用次数
        /// </summary>
        /// <param name="msg"></param>
        private async void Send(string msg)
        {
            if (!string.IsNullOrWhiteSpace(msg))
            {
                await Task.Run(() =>
                {
                    string[] strmsg = msg.Split('|');
                    for (int i = 0; i < strmsg.Length; i++)
                    {

                        try
                        {
                            lock (objSync)
                            {
                                byte[] buffer = new byte[1024 * 1024 * 3];
                                buffer = Encoding.UTF8.GetBytes("|" + strmsg[i] + "|");
                                socket.Send(buffer);
                            }
                        }
                        catch (Exception ex)
                        {
                            IsStart = false;
                            // label43.ForeColor = R;
                            // label43.Text = "看板状态：未连接";
                            ShowMsg("发送失败：" + ex.Message);
                        }
                        Thread.Sleep(10);
                        Application.DoEvents();
                    }
                });
            }
        }

        private void ShowMsg(string msg)
        {
            Invoke(new Action(() =>
            {
                if (richTextBox3.TextLength > 50000)
                {
                    richTextBox3.Clear();
                }
                string info = string.Format("{0}:{1}\r\n", DateTime.Now.ToString("G"), msg);
                richTextBox3.AppendText(info);
            }));
        }

        #endregion

        #region --------------- 处理接受过来的数据 ---------------

        bool sedbool = false;
        public void ReceiveData(string strdata)
        {
            string[] dateshuzu = new string[] { };
            dateshuzu = strdata.Split(new char[] { '+' });
            switch (dateshuzu[0])
            {
                case "0":
                    try
                    {
                        //D1204
                        KeyenceMcNet.Write("D1204", 1);
                        //D1206
                        KeyenceMcNet.Write("D1206", int.Parse(dateshuzu[1].ToString()));
                        mdb = new mdbDatas(path4);
                        DataTable table1 = mdb.Find("select * from SytemSet where ID = '1'");
                        if (table1.Rows.Count > 0)
                        {
                            string sql = "update [SytemSet] set [faults]='" + dateshuzu[1] + "'" +
                                //",Workstname ='" + checkBox6.Checked + "'" + 
                                " where [ID] = '1'";
                            var result = mdb.Change(sql);
                        }
                        mdb.CloseConnection();
                        Invoke(new Action(() =>
                        {
                            comboBox2.SelectedValue = dateshuzu[1];
                            label50.Text = dateshuzu[1];
                            Send("6+" + label54.Text + "配方切换成功");
                        }));
                    }
                    catch (Exception ex)
                    {
                        LogMsg(ex.ToString());
                        Send("6+" + label54.Text + "配方切换失败");
                    }
                    break;
                case "1":
                    try
                    {
                        if (OffLineType == 0 && loginCheck == true)
                        {
                            //MessageBox.Show("确认接收生产工单");
                            if (sedbool == true)
                            {
                                Invoke(new Action(() =>
                                {
                                    textBox6.Text = dateshuzu[1];
                                }));
                            }
                            else
                            {
                                char[] prowrd = dateshuzu[1].ToCharArray();
                                foreach (var chstr in prowrd)
                                {
                                    SendKeys.SendWait("{" + chstr + "}");
                                }
                                SendKeys.SendWait("{Enter}");
                            }
                            Invoke(new Action(() =>
                            {
                                if (textBox6.Text.Equals(dateshuzu[1]))
                                {
                                    Send("6+" + label54.Text + "生产工单发送成功");
                                }
                            }));
                        }
                    }
                    catch (Exception)
                    {
                        Send("6+" + label54.Text + "生产工单发送失败");
                    }
                    break;
            }
        }

        #endregion

        #region --------------- 打印信息 ---------------

        /// <summary>
        /// 加载打印信息
        /// </summary>
        private void PatchPZLPrinte()
        {
            try
            {
                mdb = new mdbDatas(path4);
                DataTable table1 = mdb.Find("select * from Printers where ID = 1");
                for (int i = 0; i < table1.Rows.Count; i++)
                {
                    for (int j = 0; j < table1.Columns.Count; j++)
                    {
                        for (int n = 0; n < comboBox1.Items.Count; n++)
                        {
                            if (comboBox1.GetItemText(comboBox1.Items[n]).Equals(table1.Rows[i]["Ptype"].ToString()))
                            { // 将每个选项转换为文本并存入数组
                                comboBox1.SelectedItem = table1.Rows[i]["Ptype"].ToString();
                                break;
                            }
                        }
                        // textBox21.Text = userCollection.Rows[i]["Phead"].ToString();//条码头
                        textBox25.Text = table1.Rows[i]["Ptime"].ToString();//条码时间格式
                        textBox26.Text = table1.Rows[i]["Ptail"].ToString();//条码数字
                        textBox20.Text = table1.Rows[i]["TFront"].ToString();//字体前端
                        textBox27.Text = table1.Rows[i]["Tmonarch"].ToString();//字体后端
                        textBox28.Text = table1.Rows[i]["Tlow"].ToString();//字体下面
                        //string pertow = userCollection.Rows[i]["Pertow"].ToString();//是否打印二个条码
                        string txttow = table1.Rows[i]["Txttow"].ToString();//是否打印文字
                        string plcmodel = table1.Rows[i]["Plcmodel"].ToString();//是否读取PLC型号
                        string phead = table1.Rows[i]["Phead"].ToString();//是否由PLC控制
                        label48.Text = table1.Rows[i]["Method"].ToString();
                        //if (pertow == "True") checkBox2.Checked = true;
                        if (txttow == "True") checkBox3.Checked = true;
                        if (plcmodel == "True") checkBox4.Checked = true;
                        if (phead == "True") checkBox7.Checked = true;
                    }
                }
                label62.Text = Pfunime() + textBox26.Text;
                mdb.CloseConnection();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        /// <summary>
        /// 修改打印信息
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button19_Click_1(object sender, EventArgs e)
        {
            if (textBox25.Text == String.Empty)
            {
                MessageBox.Show("当前界面内容均为必填项、请先填写完善");
                return;
            }
            try
            {
                /* PatPZLhead();
                 //FileInfo myFile1 = new FileInfo(pathtxtPZL1);
                 File.WriteAllText(pathtxtPZL1, richTextBox5.Text, Encoding.UTF8);//第二种
                 PatPZLtail();
                 File.WriteAllText(pathtxtPZL2, richTextBox6.Text, Encoding.UTF8);//第二种*/
                string method = string.Empty;
                mdb = new mdbDatas(path4);
                DataTable table1 = mdb.Find("select * from Printers where ID = 1");
                if (table1.Rows.Count > 0)
                {
                    string sql = "update [Printers] set [Ptype]='" + comboBox1.Text + "'" +
                                  ",[Ptime]='" + textBox25.Text + "',[Ptail]='" + textBox26.Text + "'" +
                                  ",[TFront]='" + textBox20.Text + "',[Tmonarch]='" + textBox27.Text + "'" +
                                  ",[Tlow]='" + textBox28.Text + "'" +
                                  ",[Txttow]='" + checkBox3.Checked + "',[Plcmodel]='" + checkBox4.Checked + "'" +
                                  ",[Method]='" + label48.Text + "',[Phead]='" + checkBox7.Checked + "'" +
                                  " where [ID] = 1";
                    var result = mdb.Change(sql);
                    if (result == true)
                    {
                        MessageBox.Show("保存成功");
                    }
                }
                label62.Text = Pfunime() + textBox26.Text;
                mdb.CloseConnection();
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }

        /// <summary>
        /// 打印机
        /// </summary>
        private void PrintZPL_Click()
        {
            while (IsRunningPZL)
            {

                // 触发通讯读
                if (plcConn == true)
                {
                    //var Read_data1 = KeyenceMcNet.ReadInt16("D1000").Content;
                    var read_data2 = KeyenceMcNet.ReadInt32("D1012").Content;      // 读取寄存器int值 
                    var read_data4 = KeyenceMcNet.ReadInt32("D1014").Content;
                    if (read_data2 == 1 && read_data4 == 0)//就去打印
                    {
                        Invoke(new Action(() =>
                        {
                            try
                            {
                                // Send a printer-specific to the printer.
                                button20_Click_1(null, null);//发送打印机指令
                                KeyenceMcNet.Write("D1014", 1);
                            }
                            catch (Exception)
                            {
                                //MessageBox.Show("发送打印机指令异常：" + ex.Message);
                                KeyenceMcNet.Write("D1014", 2);
                            }
                        }));
                        Thread.Sleep(250);
                        Application.DoEvents();
                    }
                }
            }

        }

        /// <summary>
        /// 条码前部分+日期处理
        /// </summary>
        /// <returns></returns>
        private string Pfunime()
        {
            string Ptime = textBox25.Text;
            DateTime times_Month = DateTime.Now;
            string times_Month_string0 = times_Month.Year.ToString();
            string times_Month_string1 = times_Month.Month.ToString();
            switch (times_Month_string1)
            {
                case "1": Ptime += 1; break;
                case "2": Ptime += 2; break;
                case "3": Ptime += 3; break;
                case "4": Ptime += 4; break;
                case "5": Ptime += 5; break;
                case "6": Ptime += 6; break;
                case "7": Ptime += 7; break;
                case "8": Ptime += 8; break;
                case "9": Ptime += 9; break;
                case "10": Ptime += 0; break;
                case "11": Ptime += "A"; break;
                case "12": Ptime += "B"; break;
            }
            string times_Month_string2 = times_Month.ToString("dd");
            Ptime += times_Month_string2;
            return Ptime;
        }

        /// <summary>
        /// 处理流水号数据
        /// </summary>
        private void Serialnumber()
        {
            string serialunm = textBox26.Text;
            int num;//发送一条加一
            if (int.TryParse(serialunm, out num))
            {
                SeriaPtail(num);
            }

        }

        /// <summary>
        /// 修改流水号数据
        /// </summary>
        /// <param name="num"></param>
        private void SeriaPtail(int num)
        {
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select Ptail from Printers where ID = 1");
            if (table1.Rows.Count > 0)
            {
                int numA = num + 1;
                string seriaA = numA.ToString().PadLeft(5, '0');
                string sql = "update [Printers] set [Ptail]='" + seriaA + "'" +
                              " where [ID] = 1";
                var result = mdb.Change(sql);
            }
            DataTable table2 = mdb.Find("select Ptail from Printers where ID = 1");
            for (int i = 0; i < table2.Rows.Count; i++)
            {
                for (int j = 0; j < table2.Columns.Count; j++)
                {
                    textBox26.Text = table2.Rows[i]["Ptail"].ToString();
                }
            }
            mdb.CloseConnection();
            label62.Text = Pfunime() + textBox26.Text;
        }

        /// <summary>
        /// 打印条码
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button20_Click_1(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals("无") || string.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("请选择打印机！！!");
                return;
            }
            try
            {
                int count = 1;
                int.TryParse(textBox21.Text, out count);
                for (int a = 0; a < count; a++)
                {
                    string ZPLinstructions = "";
                    string fileContent = "";
                    /*if (checkBox2.Checked)
                    {
                        fileContent = File.ReadAllText(label48.Text);
                        //双个条码指令
                        // ZPLinstructions = ZPLbarcode.ZPLdouble(textBox21.Text, Ptime, textBox26.Text, textBox20.Text, textBox27.Text, textBox28.Text);
                    }
                    else
                    {
                        // 读取文件内容到字符串中  
                        fileContent = File.ReadAllText(pathPZLsingle);
                        //单个条码指令
                        // ZPLinstructions = ZPLbarcode.ZPLsingle(textBox21.Text, Ptime, textBox26.Text, textBox20.Text, textBox27.Text, textBox28.Text);
                    }*/
                    fileContent = File.ReadAllText(label48.Text);
                    if (checkBox3.Checked)
                    {
                        ZPLinstructions = string.Format(fileContent, label62.Text, textBox20.Text, textBox27.Text, textBox28.Text);
                    }
                    else
                    {
                        ZPLinstructions = string.Format(fileContent, label62.Text);
                    }
                    RawPrinterHelper.SendStringToPrinter(comboBox1.Text, ZPLinstructions);//发送给打印机
                    Serialnumber();//将流水号+1
                }


            }
            catch (Exception ex)
            {
                string exMess = ex.Message;
                if (exMess.Equals("索引(从零开始)必须大于或等于零，且小于参数列表的大小。"))
                {
                    exMess = "请重新配置prn文件";
                }
                MessageBox.Show("发送打印机指令异常：" + exMess);
                //KeyenceMcNet.Write("D1014", 1);
            }


        }

        private void button21_Click(object sender, EventArgs e)
        {
            // PrintDialog path = new PrintDialog();
            OpenFileDialog path = new OpenFileDialog() { Filter = "Files (*.prn)|*.prn" };
            path.ShowDialog();
            this.label48.Text = path.FileName;
        }

        /// <summary>
        /// 发送打印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click_1(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(cmbInstalledPrinters.Text))
            {
                // Send a printer-specific to the printer.
                RawPrinterHelper.SendStringToPrinter(cmbInstalledPrinters.Text, this.richTextBox2.Text);
            }
        }

        #endregion

        #region --------------- PLC点位 ---------------

        //Board  
        //ID WorkID BoardName BoardCode MaxBoardCode MinBoardCode BeatBoardCode ResultBoardCode
        DataTable codesDataM;
        DataTable boardTable;

        /// <summary>
        /// 从数据库加载 Board 表格保存到对应的数组中，并为 DataGridView4,5 设置属性
        /// </summary>
        private void SYS_BOARD()
        {
            button23_Click(null, null);

            mdb = new mdbDatas(path4);
            boardTable = mdb.Find("select * from Board");

            if (boardTable.Rows.Count > 0)
            {
                sequenceNum = boardTable.AsEnumerable().Select(row => row["WorkID"].ToString()).ToArray();
                testItems = boardTable.AsEnumerable().Select(row => row["BoardName"].ToString()).ToArray();
                actualValue = boardTable.AsEnumerable().Select(row => row["BoardCode"].ToString()).ToArray();
                maxValue = boardTable.AsEnumerable().Select(row => row["MaxBoardCode"].ToString()).ToArray();
                minValue = boardTable.AsEnumerable().Select(row => row["MinBoardCode"].ToString()).ToArray();
                beat = boardTable.AsEnumerable().Select(row => row["BeatBoardCode"].ToString()).ToArray();
                testResult = boardTable.AsEnumerable().Select(row => row["ResultBoardCode"].ToString()).ToArray();
                unitName = boardTable.AsEnumerable().Select(row => row["BoardA1"].ToString()).ToArray();
                standardValue = boardTable.AsEnumerable().Select(row => row["StandardCode"].ToString()).ToArray();
            }

            mdb.CloseConnection();

            // 为 DataGridView4 添加 ButtonColumn 列，标题为操作，按钮Text默认显示为 “NO保存”
            DataGridViewButtonColumn buttonColumn = new DataGridViewButtonColumn();
            buttonColumn.HeaderText = "操作";
            buttonColumn.Text = "NO保存";
            buttonColumn.Name = "btnColNO";
            buttonColumn.DefaultCellStyle.NullValue = "NO保存";
            dataGridView4.Columns.Add(buttonColumn);

            DataGridViewButtonColumn buttonColumnOne = new DataGridViewButtonColumn();
            buttonColumnOne.HeaderText = "操作";
            buttonColumnOne.Text = "ONE保存";
            buttonColumnOne.Name = "btnColONE";
            buttonColumnOne.DefaultCellStyle.NullValue = "ONE保存";
            dataGridView4.Columns.Add(buttonColumnOne);

            // 再次创建一个新的列对象并设置其属删除
            DataGridViewButtonColumn anotherButtonColumn = new DataGridViewButtonColumn();
            anotherButtonColumn.HeaderText = "操作";
            anotherButtonColumn.Name = "btnCol2";
            anotherButtonColumn.DefaultCellStyle.NullValue = "删除";
            dataGridView4.Columns.Add(anotherButtonColumn);

            DataGridViewButtonColumn butnCo = new DataGridViewButtonColumn();
            butnCo.HeaderText = "操作";
            butnCo.Text = "保存";
            butnCo.Name = "btnCol";
            butnCo.DefaultCellStyle.NullValue = "保存";
            dataGridView5.Columns.Add(butnCo);

            DataGridViewButtonColumn anotrButCo = new DataGridViewButtonColumn();
            anotrButCo.HeaderText = "操作";
            anotrButCo.Name = "btnCol2";
            anotrButCo.DefaultCellStyle.NullValue = "删除";
            dataGridView5.Columns.Add(anotrButCo);
        }

        /// <summary>
        /// 刷新界面数据
        /// </summary>
        private void button23_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);

            // 查询Board表格， ID WorkID BoardName BoardCode MaxBoardCode MinBoardCode BeatBoardCode ResultBoardCode
            string selectSql = $"select [ID] as 编号, [WorkID] as 工位ID, [BoardName] as 测试项目的名称," +
                $"[BoardCode] as 测试项目的PLC点位, [StandardCode] as 测试项目标准值的PLC点位," +
                $"[MaxBoardCode] as 测试项目的上限PLC点位, [MinBoardCode] as 测试项目的下限PLC点位," +
                $"[ResultBoardCode] as 测试项目的结果PLC点位, [BeatBoardCode] as 测试项目的节拍PLC点位," +
                $"[BoardA1] as 单位 from Board";
            DataTable boardTable = mdb.Find(selectSql);

            // 设置 “编号” 列的属性为自动增长、只读、起始值为0
            boardTable.Columns["编号"].AutoIncrement = true;
            boardTable.Columns["编号"].ReadOnly = true;
            boardTable.Columns["编号"].AutoIncrementSeed = 0;

            // 获取 "编号" 列中的最大值
            int maxPid = 0;
            if (boardTable.Rows.Count > 0)
            {
                maxPid = boardTable.AsEnumerable().Max(row => row.Field<int>("编号"));
            }

            // 创建一个新行，设置其 "编号" 为最大值 maxPid
            DataRow newRow = boardTable.NewRow();
            newRow["编号"] = maxPid;

            // 更新 DataGridView4
            dataGridView4.DataSource = boardTable;

            codesDataM = mdb.Find("select ID as 编号 ,CName as 产品名称,TooName as 条码验证型号与工装编号,MateName as 产品编码 from Codes");
            dataGridView5.DataSource = mdb.Find("select ID as 编号 ,CName as 产品名称,TooName as 条码验证型号与工装编号,MateName as 产品编码 from Codes");

            comboBox2.DataSource = codesDataM;
            comboBox2.DisplayMember = "条码验证型号与工装编号";
            comboBox2.ValueMember = "编号";
            mdb.CloseConnection();
        }

        /// <summary>
        /// 系统设置 > PLC点位信息中，保存和删除按钮的事件处理器
        /// </summary>
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            { return; }

            // 需要记录日志的字段
            string[] logFields = { "WorkID", "BoardName", "BoardCode", "MaxBoardCode", "MinBoardCode", "BeatBoardCode", "ResultBoardCode", "BoardA1", "StandardCode" };
            // 定义字段与别名的映射
            Dictionary<string, string> fieldAliases = new Dictionary<string, string>
            {
                { "WorkID", "工位序号" },
                { "BoardName", "测试项" },
                { "BoardCode", "实际值点位" },
                { "MaxBoardCode", "上限值点位" },
                { "MinBoardCode", "下限值点位" },
                { "BeatBoardCode", "节拍点位" },
                { "ResultBoardCode", "测试结果" },
                { "BoardA1", "单位" },
                { "StandardCode", "标准值点位" },
            };

            // 从DataGridView 删除选定行
            DeleteRowFromDataGridView<int>(dataGridView4, e, "Board", "ID", logFields, 3, path4, "btnClo2", fieldAliases, "PLC点位数据");

            // NO保存
            if (dataGridView4.Columns[e.ColumnIndex].Name == "btnColNO")
            {
                //Board  
                //ID WorkID BoardName BoardCode MaxBoardCode MinBoardCode BeatBoardCode ResultBoardCode 
                //说明点击的列是DataGridViewButtonColumn列
                DataGridViewColumn column = dataGridView4.Columns[e.ColumnIndex];
                string pid = this.dataGridView4.Rows[e.RowIndex].Cells[3].Value.ToString();
                string workID = this.dataGridView4.Rows[e.RowIndex].Cells[4].Value.ToString();
                string boardName = this.dataGridView4.Rows[e.RowIndex].Cells[5].Value.ToString();
                string boardCode = this.dataGridView4.Rows[e.RowIndex].Cells[6].Value.ToString();
                string stanCode = CodeNum.NOCodes(this.dataGridView4.Rows[e.RowIndex].Cells[7].Value.ToString());
                string maxBoardCode = CodeNum.NOCodes(this.dataGridView4.Rows[e.RowIndex].Cells[8].Value.ToString());
                string minBoardCode = CodeNum.NOCodes(this.dataGridView4.Rows[e.RowIndex].Cells[9].Value.ToString());
                string resultBoardCode = CodeNum.NOCodes(this.dataGridView4.Rows[e.RowIndex].Cells[10].Value.ToString());
                string beatBoardCode = CodeNum.NOCodes(this.dataGridView4.Rows[e.RowIndex].Cells[11].Value.ToString());
                string beatBoardA1 = CodeNum.DaweiCodes(this.dataGridView4.Rows[e.RowIndex].Cells[12].Value.ToString());

                Barddrefom(pid, workID, stanCode, boardName, boardCode, maxBoardCode, minBoardCode, resultBoardCode, beatBoardCode, beatBoardA1);
                button23_Click(null, null);
            }
            // ONE保存
            if (dataGridView4.Columns[e.ColumnIndex].Name == "btnColONE")
            {
                DataGridViewColumn column = dataGridView4.Columns[e.ColumnIndex];
                string pid = this.dataGridView4.Rows[e.RowIndex].Cells[3].Value.ToString();
                string workID = this.dataGridView4.Rows[e.RowIndex].Cells[4].Value.ToString();
                string boardName = this.dataGridView4.Rows[e.RowIndex].Cells[5].Value.ToString();
                string boardCode = this.dataGridView4.Rows[e.RowIndex].Cells[6].Value.ToString();
                string stanCode = CodeNum.ONECodes(this.dataGridView4.Rows[e.RowIndex].Cells[7].Value.ToString());
                string maxBoardCode = CodeNum.ONECodes(this.dataGridView4.Rows[e.RowIndex].Cells[8].Value.ToString());
                string minBoardCode = CodeNum.ONECodes(this.dataGridView4.Rows[e.RowIndex].Cells[9].Value.ToString());
                string resultBoardCode = CodeNum.ONECodes(this.dataGridView4.Rows[e.RowIndex].Cells[10].Value.ToString());
                string beatBoardCode = CodeNum.ONECodes(this.dataGridView4.Rows[e.RowIndex].Cells[11].Value.ToString());
                string beatBoardA1 = CodeNum.DaweiCodes(this.dataGridView4.Rows[e.RowIndex].Cells[12].Value.ToString());

                Barddrefom(pid, workID, stanCode, boardName, boardCode, maxBoardCode, minBoardCode, resultBoardCode, beatBoardCode, beatBoardA1);
                button23_Click(null, null);
            }
        }

        /// <summary>
        /// 删除 DataGridView 控件中的某一行，并记录删除的数据
        /// </summary>
        private void DeleteRowFromDataGridView<T>(DataGridView dataGridView, DataGridViewCellEventArgs e,
            string tableName, string primaryKeyField, string[] logFields, int primaryKeyColumnIndex,
            string connectionString, string btnName, Dictionary<string, string> fieldAliases, string viewName)
        {
            // 检查选中的行是否为空
            if (dataGridView.Rows[e.RowIndex].Cells[primaryKeyColumnIndex].Value == null ||
                string.IsNullOrWhiteSpace(dataGridView.Rows[e.RowIndex].Cells[primaryKeyColumnIndex].Value.ToString()))
            {
                //MessageBox.Show("选中的行无效或编号为空，无法删除！");
                return;
            }

            if (dataGridView.Columns[e.ColumnIndex].Name == "btnCol2")
            {
                // 获取选中行的主键值，并将其转换为泛型类型
                T primaryKeyValue = (T)dataGridView.Rows[e.RowIndex].Cells[primaryKeyColumnIndex].Value;

                // 根据主键类型创建 SQL 语句
                string primaryKeyCondition = primaryKeyValue is string ? $"'{primaryKeyValue}'" : primaryKeyValue.ToString();

                // 查询要删除的行数据
                mdb = new mdbDatas(connectionString);
                string selectSql = $"SELECT * FROM {tableName} WHERE [{primaryKeyField}] = {primaryKeyCondition}";
                DataTable rowTable = mdb.Find(selectSql);

                if (rowTable.Rows.Count > 0)
                {
                    DataRow row = rowTable.Rows[0];

                    // 生成日志详细数据
                    string logDetail = $"编号: {row[primaryKeyField]}";
                    foreach (var field in logFields)
                    {
                        //logDetail += $" | {field}: {row[field]}";
                        if (fieldAliases.TryGetValue(field, out string alias))
                        {
                            logDetail += $" | {alias}：{row[field]}";
                        }
                        else
                        {
                            // 如果映射中没有找到别名，就使用字段名
                            logDetail += $" | {field}：{row[field]}";
                        }
                    }

                    // 执行删除操作
                    string deleteSql = $"DELETE FROM {tableName} WHERE [{primaryKeyField}] = {primaryKeyCondition}";
                    bool bl = mdb.Del(deleteSql);

                    if (bl)
                    {
                        MessageBox.Show("删除成功");
                        loggerConfig.Trace($"【{viewName}删除】\n" +
                                           $"成功删除第{primaryKeyCondition}行, 该行的详细数据: \n{logDetail}");

                        button23_Click(null, null);
                        button18_Click(null, null);
                    }
                }

                mdb.CloseConnection();
            }
        }

        /// <summary>
        /// 更新或新增 DataGridView4 的数据
        /// </summary>
        private void Barddrefom(string pid, string workID, string stanCode, string boardName, string boardCode, string maxBoardCode, string minBoardCode, string resultBoardCode, string beatBoardCode, string beatBoardA1)
        {
            mdb = new mdbDatas(path4);
            string selectSql = $"select * from Board where [ID] = {pid}";
            DataTable table1 = mdb.Find(selectSql);

            // 修改
            if (table1.Rows.Count > 0)
            {
                DataRow row = table1.Rows[0];

                // 记录详细数据到日志
                string logDetail = $"编号：{row["ID"]} | 工位序号：{row["WorkID"]} | 测试项：{row["BoardName"]} | " +
                                   $"实际值点位：{row["BoardCode"]} | 上限值点位：{row["MaxBoardCode"]} | " +
                                   $"下限值点位：{row["MinBoardCode"]} | 节拍点位：{row["BeatBoardCode"]} | " +
                                   $"测试结果点位：{row["ResultBoardCode"]} | 单位：{row["BoardA1"]} | " +
                                   $"标准值点位：{row["StandardCode"]}";

                if (table1.Rows.Count > 0)
                {
                    string sql = $"update [Board] set [WorkID] = '{workID}', [BoardName] = '{boardName}', " +
                                 $"[BoardCode] = '{boardCode}', [MaxBoardCode] = '{maxBoardCode}', [MinBoardCode] = '{minBoardCode}', " +
                                 $"[BeatBoardCode] = '{beatBoardCode}', [ResultBoardCode] = '{resultBoardCode}', [BoardA1] = '{beatBoardA1}', " +
                                 $"[StandardCode] = '{stanCode}' where [ID] = {pid}";

                    var result = mdb.Change(sql);
                    if (result == true)
                    {
                        MessageBox.Show("修改成功");
                        string modifyInfo = $"【点位数据修改成功】\n修改前的详细信息：\n{logDetail}\n修改后的详细信息：\n编号：{pid} | 工位序号：{workID} | 测试项：{boardName} | " +
                                            $"实际值点位：{boardCode} | 上限值点位：{maxBoardCode} | 下限值点位：{minBoardCode} | 节拍点位：{beatBoardCode} | " +
                                            $"测试结果点位：{resultBoardCode} | 单位：{beatBoardA1} | 标准值点位：{stanCode} ";
                        loggerConfig.Trace(modifyInfo);
                    }
                }

            }
            // 新增
            else
            {
                string sql = "insert into Board ([ID],[WorkID],[StandardCode],[BoardName],[BoardCode],[MaxBoardCode],[MinBoardCode],[BeatBoardCode],[ResultBoardCode],[BoardA1]) values ("
                    + pid + ",'" + workID + "','" + stanCode + "','" + boardName + "','" + boardCode + "','" + maxBoardCode + "','" + minBoardCode + "','" + beatBoardCode + "','" + resultBoardCode + "','" + beatBoardA1 + "')";

                bool result = mdb.Add(sql.ToString());
                if (result == true)
                {
                    MessageBox.Show("新增成功");
                    loggerConfig.Trace($"【点位数据新增成功】\n新增详情：\n" +
                                       $"编号：{pid} | 工位序号：{workID} | 测试项：{boardName} | " +
                                       $"实际值点位：{boardCode} | 上限值点位：{maxBoardCode} | 下限值点位：{minBoardCode} | " +
                                       $"节拍点位：{beatBoardCode} | 测试结果：{resultBoardCode} | 单位：{beatBoardA1} | 标准值点位：{stanCode} ");
                }
            }

            mdb.CloseConnection();
        }

        /// <summary>
        /// DataGridView5 中 保存和删除按钮的事件处理器
        /// </summary>
        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
            { return; }

            // 需要记录日志的字段
            string[] logFields = { "CName", "TooName" };

            // 定义字段与别名的映射
            Dictionary<string, string> fieldAliases = new Dictionary<string, string>
            {
                { "CName", "产品名称" },
                { "TooName", "条码规则" },
            };

            // 删除
            DeleteRowFromDataGridView<string>(dataGridView5, e, "Codes", "ID", logFields, 2, path4, "btnClo2", fieldAliases, "配方信息");

            // 保存
            if (dataGridView5.Columns[e.ColumnIndex].Name == "btnCol")
            {
                //DataGridViewColumn column = dataGridView5.Columns[e.ColumnIndex];
                string pid = this.dataGridView5.Rows[e.RowIndex].Cells[2].Value.ToString();         // 编号
                string funmae = this.dataGridView5.Rows[e.RowIndex].Cells[3].Value.ToString();      // 产品名称
                string tooname = this.dataGridView5.Rows[e.RowIndex].Cells[4].Value.ToString();     // 条码验证型号与工装编号
                string mateName = this.dataGridView5.Rows[e.RowIndex].Cells[5].Value.ToString();    // 产品编码

                if (string.IsNullOrWhiteSpace(pid))
                {
                    MessageBox.Show("编号不能为空！");
                    return;
                }

                mdb = new mdbDatas(path4);
                string selectSql = $" select * from Codes where [ID] = '{pid}' ";
                DataTable table1 = mdb.Find(selectSql);

                // 修改
                if (table1.Rows.Count > 0)
                {
                    DataRow row = table1.Rows[0];

                    // 记录修改前的详细数据
                    string logDetail = $"编号：{row["ID"]} | 产品名称：{row["CName"]} | 条码验证型号与工装编号：{row["TooName"]} | 产品编码：{row["MateName"]} ";

                    string updateSql = $"update [Codes] set [CName]='{funmae}', [TooName] ='{tooname}', [MateName] ='{mateName}' where [ID] = '{pid}'";
                    var result = mdb.Change(updateSql);

                    if (result == true)
                    {
                        MessageBox.Show("修改成功");
                        string modifyInfo = $"【产品信息修改成功】\n" +
                                            $"修改前的详细信息：\n{logDetail}\n" +
                                            $"修改后的详细信息：\n编号：{pid} | 产品名称：{funmae} | 条码验证型号与工装编号：{tooname} | 产品编码：{mateName}";
                        loggerConfig.Trace(modifyInfo);
                    }
                }

                // 新增
                else
                {
                    string sql = $" insert into Codes ([ID],[CName],[TooName],[MateName]) values('{pid}','{funmae}','{tooname}','{mateName}') ";
                    bool result = mdb.Add(sql.ToString());
                    if (result == true)
                    {
                        MessageBox.Show("新增成功");
                        loggerConfig.Trace($"【产品信息新增成功】\n" +
                                           $"新增详情：\n" +
                                           $"编号：{pid} | 产品名称：{funmae} | 条码验证型号与工装编号：{tooname} | 产品编码：{mateName}");
                    }
                }

                mdb.CloseConnection();
            }
        }

        #endregion

        #region --------------- 读卡器 ---------------

        RfidReader Reader = new RfidReader();
        List<UserInfoEntity> userInfoEntities;
        Task taskProcess_UserID = null;
        string PuserUID = string.Empty;
        bool isReaderOpen = false;

        public void SearchPort()
        {
            string[] ports = SerialPort.GetPortNames();

            cmbShowPort.Items.Clear();
            cmbShowPort.Text = null;

            foreach (string port in ports)
            {
                cmbShowPort.Items.Add(port);
            }

            if (ports.Length > 0)
            {
                cmbShowPort.Text = ports[0];
            }
            //else
            //{
            //    MessageBox.Show("没有发现可用端口");
            //}
        }

        /// <summary>
        /// 搜索读卡器端口
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SearchPort_Click(object sender, EventArgs e)
        {
            SearchPort();
        }

        /// <summary>
        /// 连接读卡器
        /// </summary>
        private void OpenReader_Click(object sender, EventArgs e)
        {
            try
            {
                Reader.DisConnect();

                // 判断读卡器当前连接状态
                if (isReaderOpen)
                {
                    btnOpenReader.Enabled = false;
                    isReaderOpen = false;
                    btnOpenReader.Text = resources.GetString("readCard_Open");
                    Task.WaitAll(taskProcess_UserID);
                    Thread.Sleep(1000);
                    btnOpenReader.Enabled = true;
                }
                else
                {
                    // 连接读卡器
                    if (cmbShowPort.Text.Length < 1 || tbxReaderDeviceID.Text.Length < 1)
                    {
                        lblReaderState.Text = (resources.GetString("posType"));
                        lblReaderState.ForeColor = Color.Red;
                        return;
                    }
                    else
                    {
                        bool flg = Reader.Connect(cmbShowPort.Text, 9600);
                        if (flg == true)
                        {
                            lblReaderState.Text = "成功连接读卡器";
                            lblReaderState.ForeColor = Color.Green;
                            isReaderOpen = true;
                            btnOpenReader.Text = resources.GetString("readCard_Close");

                            taskProcess_UserID = new Task(Process_UserID);
                            taskProcess_UserID.Start();
                        }
                        else
                        {
                            lblReaderState.Text = (resources.GetString("portNG"));
                            lblReaderState.ForeColor = Color.Red;
                            return;
                        }
                    }
                }
            }
            catch
            {
                lblReaderState.Text = (resources.GetString("portType"));
                lblReaderState.ForeColor = Color.Red;
                return;
            }

        }

        private void Process_UserID()
        {
            while (isReaderOpen)
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
            try
            {
                int status;
                byte[] type = new byte[2];
                byte[] id = new byte[4];

                Reader.Cmd = Cmd.M1_ReadId;//读卡号命令
                Reader.Addr = Convert.ToByte(tbxReaderDeviceID.Text, 16);//读写器地址,设备号
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

                    string userid = byteToHexStr(id, 4);
                    if (userid.Length > 0)
                    {
                        tbxBrandID.Text = userid;
                        PuserUID = userid;

                        //查询用户
                        //根据userid 读取账号密码
                        List<UserInfoEntity> list = userInfoEntities.Where(x => x.cardID == PuserUID).ToList();
                        if (list.Count > 0)
                        {
                            if (list.Count > 1)
                            {
                                KeyenceMcNet.Write("D18040", -1);
                            }
                            foreach (var v in list)
                            {
                                int Paccess = 0;
                                if (v.Uuser.Length > 0 && v.Upwd.Length > 0)
                                {
                                    if (v.Utype == "ADM")
                                    {
                                        Paccess = 5;
                                    }
                                    else if (v.Utype == "PE")
                                    {
                                        Paccess = 2;
                                    }
                                    else if (v.Utype == "QE")
                                    {
                                        Paccess = 4;
                                    }
                                    else if (v.Utype == "ME")
                                    {
                                        Paccess = 3;
                                    }
                                    else if (v.Utype == "OP")
                                    {
                                        Paccess = 1;
                                    }
                                }
                                lblPlcAccess.Text = Paccess.ToString();
                                KeyenceMcNet.Write("D18040", Paccess);
                                loggerAccount.Trace($"【PLC触摸屏当前权限信息】\n 工号：{v.Uuser} | 姓名：{v.userName} | 权限：{v.Utype}");
                            }
                        }
                        else
                        {
                            KeyenceMcNet.Write("D18040", -1);
                        }

                        PuserUID = string.Empty;
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
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

        #endregion

        #region --------------- PLC测试 ---------------

        /// <summary>
        /// 读取
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button27_Click(object sender, EventArgs e)
        {
            textBox31.Text = KeyenceMcNet.ReadInt32(textBox23.Text).Content.ToString();
            MessageBox.Show("执行完成！");
        }
        private void button28_Click(object sender, EventArgs e)
        {
            KeyenceMcNet.Write(textBox29.Text, int.Parse(textBox36.Text));
            MessageBox.Show("执行完成！");
        }

        #endregion

        private void dataGridViewDynamic1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            try
            {
                //示例：根据第4列的状态0,1,2 显示不同的颜色
                string status = dataGridViewDynamic1.Rows[e.RowIndex].Cells[7].Value.ToString();
                switch (status)
                {
                    case "NG":
                        dataGridViewDynamic1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                        break;
                    case "OK":
                        //dataGridViewDynamic1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Green;
                        break;
                }
            }
            catch
            {

            }
        }

    }
}