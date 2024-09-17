using MesDatas;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BulletinBoard
{
    public partial class Form1 : Form
    {
        public static string path4 = System.AppDomain.CurrentDomain.BaseDirectory + "ProdModel.mdb";
        mdbDatas mdb =null;
        mdbDatas mdbABC = new mdbDatas();
       // mdbDatas mdb1 = new mdbDatas();
        private List<Socket> ClientSockets;
        private BindingList<string> ClientIPPorts;
        private Socket socketWatch;
        private Socket socketSend;//发送
        private bool IsStart;
        private Action<string> ShowMsgAction;
        private Action UpdateListViewDataAction;
        DataTable stationTable;//机台
        DataTable productTable;
        public Form1()
        {
            //StopProcess();
            InitializeComponent();
            ProductUpdeteconn();
            ShowMsgAction += new Action<string>(ShowMsg);
            UpdateListViewDataAction += new Action(UpdateListViewData);
            ClientSockets = new List<Socket>();
            ClientIPPorts = new BindingList<string>();
            
            // BindingSource bs = new BindingSource(); //实例化对象
            //  bs.DataSource = ClientIPPorts; //将List<T>绑定到BindingSource中
             // listBox1.DataSource = bs; //将BindingSource绑定至最终的使用者（控件）的数据源

        }
       
        /// <summary>
        /// 窗体加载时
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            SYS_BOARD();
            SystBoardnt();//获取产线基本设置
            button8_Click(null, null);
            //mdb = new mdbDatas();
            string conn = label10.Text + "\\" + DateTime.Now.ToString("Y") + "产线数据.mdb";
            if (mdbABC.mdbDatesconn(conn) == false)
            {
                ProductLines();//初始化数据库
                // mdb.OpenConnction();
                mdbABC.OpenConnction();
                /// ShowMsg("打开数据库");
            }
            mdbABC.OpenConnction();
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //LineModel_Write();
            //IsRunning = false;
            //StopTools();
            //dataGridView1.updateData -= ProcCompleteData;
            //SaveParam();
            
            Environment.Exit(0);
        }
        /// <summary>
        /// 修改客服端连接信息
        /// </summary>
        private static void ProductUpdeteconn()
        {
           // mdbDatas mdbDa = new mdbDatas(path4);//conndnew connt
            string sql = "update  [product] set IP ='  ' " +
                ",conndnew='" +"不详"+ "'" +
                ",connt='断开'  ";
           // bool resultC = mdbDa.Add(sql.ToString());
           // mdbDa.CloseConnection();
        }

        public bool IsRunningCheckCard = true;
        public int accesscard;
        public int access;
        /// <summary 
        /// 窗体加载后
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Shown(object sender, EventArgs e)
        {
            Thread thread = new Thread(ConnSocket);
            thread.IsBackground = true;
            thread.Start();
        }
        #region--------服务器设置-------
        /// <summary>
        /// 加载服务器数据
        /// </summary>
        private void SystBoardnt()
        {
            Prodes();
            mdb = new mdbDatas(path4);
            //bool bl = mdb.Del(" DELETE FROM  [ModPros] WHERE [A2] = '" + 1 + "'");
            DataTable table1 = mdb.Find("select * from Bulletins where ID = '1'");

            for (int i = 0; i < table1.Rows.Count; i++)
            {
                for (int j = 0; j < table1.Columns.Count; j++)
                {
                    textBox7.Text = tbx_IP.Text = table1.Rows[i]["IP"].ToString();
                    textBox6.Text = tbx_port.Text = table1.Rows[i]["Port"].ToString();
                    textBox1.Text = table1.Rows[i]["BaseName"].ToString();
                    textBox2.Text = table1.Rows[i]["WorkshopName"].ToString();
                    textBox3.Text = table1.Rows[i]["ProductLineName"].ToString();
                    textBox4.Text = table1.Rows[i]["ProductLineDescription"].ToString();
                    comboBox1.SelectedItem = table1.Rows[i]["ProductLlineAttributes"].ToString();
                    label10.Text = table1.Rows[i]["Posits"].ToString();
                    textBox8.Text = table1.Rows[i]["WorkName"].ToString();
                    string fioo= table1.Rows[i]["FinishedName"].ToString();
                    if (fioo == "True")
                    {
                        checkBox1.Checked = true;
                    }
                    textBox9.Text = table1.Rows[i]["DegreesSev"].ToString();
                }
            }     
            mdb.CloseConnection();
            //加载工位
        }
        /// <summary>
        /// 加载工位
        /// </summary>
        private void Prodes()
        {
            mdb = new mdbDatas(path4);
            string strStat = "";
            stationTable = mdb.Find("select * from Model");
           // comboBox2.DataSource = stationTable;
           // comboBox2.DisplayMember = "Mname";
            //comboBox2.ValueMember = "ID";//displaymember
            foreach (DataRow row in stationTable.Rows)
            {
                strStat += "{" + row[0] + "}\t";
            }
            label7.Text = "总共：" + stationTable.Rows.Count;
            textBox5.Text = strStat;
            mdb.CloseConnection();
        }

        /// <summary>
        /// 保存服务器配置设置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_2(object sender, EventArgs e)
        {
            //Bulletins  表
            //ID
            //BaseName
            //WorkshopName
            //ProductLineName
            //ProductLineDescription
            //ProductLlineAttributes
            //Posits
            if (textBox1.Text == String.Empty || textBox2.Text == String.Empty || textBox3.Text == String.Empty
                || textBox6.Text == String.Empty || textBox7.Text == String.Empty)
            {
                MessageBox.Show("当前界面内容均为必填项、请先填写完善");
                return;
            }
            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from Bulletins where ID = '1'");
            if (table1.Rows.Count > 0)
            {
                string sql = "update [Bulletins] set [BaseName]='" + textBox1.Text + "',[WorkshopName]='" + textBox2.Text + "'" +
                              ",[ProductLineName]='" + textBox3.Text + "',[ProductLineDescription]='" + textBox4.Text + "'" +
                              ",[ProductLlineAttributes]='" + comboBox1.Text + "',[IP]='" + textBox7.Text + "'" +
                              ",[Port]='" + textBox6.Text  + "'" + ",[Posits]='" + label10.Text + "'" +
                              " where [ID] = '1'";
                var result = mdb.Change(sql);
                if (result)
                {
                     tbx_IP.Text = textBox7.Text;
                    tbx_port.Text = textBox6.Text;
                    MessageBox.Show("保存成功");
                }
            }
            mdb.CloseConnection();
        }
        /// <summary>
        /// 清空工位表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);
            bool bl = mdb.Del(" DELETE FROM  [product] ");
            bool b2 = mdb.Del(" DELETE FROM  [Model] ");
            bool b3 = mdb.Del(" DELETE FROM  [ModPros] ");
            if (bl == true)
            {
                textBox5.Text = "";
                MessageBox.Show("清空成功");
            }
            label7.Text = "总共：0";
            mdb.CloseConnection();
        }
        /// <summary>
        /// 修改最后工位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            //FinishedName
            //WorkName
            //DegreesSev
            if ( textBox9.Text == String.Empty)
            {
                MessageBox.Show("当前界面内容均为必填项、请先填写完善");
                return;
            }

            mdb = new mdbDatas(path4);
            DataTable table1 = mdb.Find("select * from Bulletins where ID = '1'");
            if (table1.Rows.Count > 0)
            {
                string sql = "update [Bulletins] set [WorkName]='" + textBox8.Text + "'"+",[FinishedName]='" + checkBox1.Checked + "'" +
                              ",[DegreesSev]='" + textBox9.Text +  "'" + " where [ID] = '1'";
                var result = mdb.Change(sql);
                if (result)
                {
                    tbx_IP.Text = textBox7.Text;
                    tbx_port.Text = textBox6.Text;
                    MessageBox.Show("保存成功");
                }
            }
            mdb.CloseConnection();
        }
        #endregion

        #region-----产线数据库-----------------------
        /// <summary>
        /// 重新生成数据库
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            //DateTime times_Month = DateTime.Now;
            //string times_Month_string0 = times_Month.Year.ToString();
            //string times_Month_string1 = times_Month.Month.ToString();
            string conn = label10.Text + "\\" + DateTime.Now.ToString("Y") + "产线数据.mdb";
            mdb = new mdbDatas();
            if (mdb.mdbDatescomm(conn) == false)
            {
                ProductLines();
                MessageBox.Show("数据库初始成功！");
            }
            else
            {
                MessageBox.Show("该数据库已经存在，请删除数据库在初始化！");
            }
            mdb.CloseConnection();
        }
        /// <summary>
        /// 生成产线数据库
        /// </summary>
        private void ProductLines()
        {
            string conn = label10.Text + "\\" + DateTime.Now.ToString("Y") + "产线数据.mdb";
            mdbDatas.CreateAccessDatabase(conn);
            //创建产线信息表
            StringBuilder str = new StringBuilder(" 基地名称,车间名称,产线名称,产线工位数量,产线描述,产线属性");
            ArrayList arrayList = new ArrayList();
            object[] obj = new object[] { "基地名称", "车间名称", "产线名称", "产线工位数量", "产线描述", "产线属性" };
            arrayList.AddRange(obj);
            if (stationTable.Rows.Count > 0)
            {
                object[] obj1 = new object[stationTable.Rows.Count];
                for (int i = 0; i < stationTable.Rows.Count; i++)
                {
                    str.Append(",");
                    int count = (i + 1);
                    obj1[i] = "工位" +count;
                    str.Append("工位" +count);
                }
                arrayList.AddRange(obj1);
            }
            mdbDatas.CreateMDBTable(conn, "产线信息", arrayList);

            //创建故障信息表
            ArrayList arrayList1 = new ArrayList();
            object[] obje1 = new object[] { "故障发生工位", "机台名称", "故障类型", "故障描述", "发生时间", "结束时间", "更新标识" };
            arrayList1.AddRange(obje1);
            mdbDatas.CreateMDBTable(conn, "故障信息", arrayList1);

            //创建生产信息表
            ArrayList arrayList2 = new ArrayList();
            object[] obje2 = new object[] {  "工位名称", "当前工单号","工装编号","产品编码","当前配方名称", "产品条码", "操作人员", "测试时间",
                "测试结果","测试节拍","测试项名称","测试项上限","测试项下限","测试项的实际值", "更新标识" };
            arrayList2.AddRange(obje2);
            mdbDatas.CreateMDBTable(conn, "生产信息", arrayList2);//new System.Collections.ArrayList(new object[] { "产品", "条码", "测试人", "测试时间","测试结果", "PLC配方", "文件版本","软件版本"}));


            //创建统计信息表
            ArrayList arrayList3 = new ArrayList();
            object[] obje3 = new object[] { "工单号", "成品名称", "工单数量", "完成数量", "完成率", "合格率",
                "整线节拍","线平衡","OEE","直通率","更新时间","更新标识" };
            arrayList3.AddRange(obje3);
            mdbDatas.CreateMDBTable(conn, "统计信息", arrayList3);


            //创建易损件信息表
            ArrayList arrayList4 = new ArrayList();
            object[] obje4 = new object[] { "易损件所在工位", "机台名称", "易损件所在位置", "易损件名称", "易损件理论使用次数",
                "易损件已使用次数","易损件剩余使用次数" };
            arrayList4.AddRange(obje4);
            mdbDatas.CreateMDBTable(conn, "易损件信息", arrayList4);
            //初始化产线信息数据
            mdb = new mdbDatas(conn);
            DataTable table1 = mdb.Find("select * from 产线信息 where ID = 1");
            if (table1.Rows.Count == 0)
            {
                // DateTime now = DateTime.Now;
                StringBuilder str1 = new StringBuilder();
                // str1.Append("'" + "1" + "',");
                str1.Append("'" + textBox1.Text + "',");
                str1.Append("'" + textBox2.Text + "',");
                str1.Append("'" + textBox3.Text + "',");
                str1.Append("'" + stationTable.Rows.Count + "',");
                str1.Append("'" + textBox4.Text + "',");
                str1.Append("'" + comboBox1.Text + "'");
                if (stationTable.Rows.Count > 0)
                {
                    for (int i = 0; i < stationTable.Rows.Count; i++)
                    {
                        str1.Append(",");
                        string rst = "  ";
                        rst = stationTable.Rows[i][0].ToString();
                        str1.Append("'" + rst + "'");

                    }
                }

                string sql = "insert into 产线信息 (" + str + ") values (" + str1 + ")";
                bool result = mdb.Add(sql.ToString());
                if (result)
                {
                    ShowMsg("数据库生成成功!");
                }
            }
            mdb.CloseConnection();
        }
        #endregion

        #region--------添加生产数据-------------
        /// <summary>
        /// 将客服端发来信息整理保存数据库
        /// </summary>
        /// <param name="strdata"></param>
        private void Datapress(string strdata,string SIP)
        {
            Invoke(new Action(() =>
            {
                string[] bulletindata = strdata.Split('|');
                foreach (var balles in bulletindata)
                {
                    string[] dateshuzu = new string[] { };
                    dateshuzu = balles.Split(new char[] { '+' });
                    for (int i = 0; i < dateshuzu.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(dateshuzu[i]))
                        {
                            dateshuzu[i] = "";
                        }
                    }
                    /*mdb = new mdbDatas();
                    string conn = label10.Text + "\\" + DateTime.Now.ToString("Y") + "产线数据.mdb";
                    if (mdb.mdbDatesconn(conn) == false)
                    {
                        ProductLines();//初始化数据库
                                       // mdb.OpenConnction();
                        mdb.OpenConnction();
                       /// ShowMsg("打开数据库");
                    }*/
                    mdbABC.OpenConnction();
                    /// 1+故障所在工位+机台名称+ 故障状态+故障的描述+触发故障的开始时间 
                    /// 1+故障所在工位+机台名称+ 故障状态+故障的描述+触发故障的结束时间
                    /// 2+工位名称+当前工单号+产品条码+操作人员+则式时间+则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值
                    /// 3+工位+工单数量+完成数量+完成率+合格率+整体节拍+生产产品数量（总数）+ 工序时间+利用时间+负荷时间
                    /// 4+易损件所在工位+机台名称+ 易损件所在位置+易损件名称+易损件理论使用次数易损件已使用次数

                    switch (dateshuzu[0])
                    {
                        case "0":
                            DataRow[] rows = stationTable.Select("Model = '" + dateshuzu[1] + "'");
                            //conntatdt
                            if (rows.Length == 0)
                            {
                                try
                                {
                                    int stationCount = stationTable.Rows.Count + 1;
                                    mdbDatas mdbDa = new mdbDatas(path4);//conndnew connt
                                    string sqlC = "insert into [Model] ([Model],[Mname]) values ('"
                                    + dateshuzu[1] + "','" + "工位" + stationCount + "')";
                                    bool resultC = mdbDa.Add(sqlC.ToString());
                                    mdbDa.CloseConnection();
                                    SystBoardnt(); //加载工位
                                    string sqlA = "ALTER TABLE [产线信息] ADD " + "工位" + stationCount + "  varchar(200)";
                                    bool resultA = mdbABC.Add(sqlA);
                                    string sqlB = "update  [产线信息] set 产线工位数量 ='" + stationCount +
                                    "',[工位" + stationCount + "]='" + dateshuzu[1] + "'where ID=1 ";
                                    bool resultB = mdbABC.Add(sqlB.ToString());
                                }
                                catch { }
                                //button8_Click(null, null);//刷新状态
                            }
                            break;
                        case "1"://故障信息
                            GweiNameID(dateshuzu);
                            NewMethod1(dateshuzu);
                            break;
                        case "2"://生产信息
                            GweiNameID(dateshuzu);
                            NewMethod2(dateshuzu);
                            break;
                        case "3"://统计信息ModPros
                            NewMethod3(dateshuzu);
                            break;
                        case "4"://易损件信息
                            GweiNameID(dateshuzu);
                            NewMethod4(dateshuzu);
                            break;
                        case "5":
                            StatName(SIP, dateshuzu);
                            break;
                    }
                }
               //mdb.CloseConnection();
            }));
        }
        /// <summary>
        /// 将名称修改工位ID
        /// </summary>
        /// <param name="dateshuzu"></param>
        private void GweiNameID(string[] dateshuzu)
        {
            if (checkBox1.Checked)
            {//[Model] ([Model],[Mname]
                DataRow[] rows = stationTable.Select("Model = '" + dateshuzu[1] + "'");
                if (rows.Length > 0)
                {
                    dateshuzu[1] = rows[0]["Mname"].ToString();
                }
            }
        }
        //DataTable productTable
        /// <summary>
        /// 添加机台名称
        /// </summary>
        /// <param name="SIP"></param>
        /// <param name="dateshuzu"></param>
        private void StatName(string SIP, string[] dateshuzu)
        {
            DataRow[] rows = productTable.Select("名称 = '" + dateshuzu[1] + "'");
            if (rows.Length == 0)
            {
                mdbDatas mdbDa = new mdbDatas(path4);//conndnew connt
                string sqlC = "insert into [product] ([Mname],[IP],[conndnew],[connt]) values ('"
                + dateshuzu[1] + "','" + SIP + "','不详','断开')";
                bool resultC = mdbDa.Add(sqlC.ToString());
                mdbDa.CloseConnection();
                productTable.Rows.Add(dateshuzu[1], SIP, DateTime.Now.ToString(), "成功");

                //button8_Click(null, null);
            }
            else
            {
               // productTable[dateshuzu[1]].;
                    // 名称, IP as IP,conndnew as 时间 ,connt as 状态
                    if (rows[0]["名称"].Equals(dateshuzu[1]))
                    {
                        rows[0]["IP"] = SIP;
                        rows[0]["时间"] = DateTime.Now;
                        rows[0]["状态"] = "成功";
                    }
               // button8_Click(null, null);
            }

            dataGridView1.DataSource = productTable;

           /* for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                string str = dataGridView1.Rows[i].Cells["状态"].Value.ToString();
                if (str.Equals("断开"))
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                }
            }*/
        }

        /// <summary>
        /// 易损件信息
        /// </summary>
        /// <param name="dateshuzu"></param>
        private void NewMethod4(string[] dateshuzu)
        {
            /// 4+易损件所在工位+机台名称+ 易损件所在位置+易损件名称+易损件理论使用次数+易损件已使用次数
            /// 易损件所在工位", "机台名称", "易损件所在位置", "易损件名称", "易损件理论使用次数",
            // "易损件已使用次数","易损件剩余使用次数"
            int count4_1 = 0;//易损件理论使用次数
            int count4_2 = 0;//易损件已使用次数
            int count4_3 = 0;//易损件剩余使用次数
            if (int.TryParse(dateshuzu[5], out count4_1) && int.TryParse(dateshuzu[6], out count4_2))
            {
                count4_3 = count4_1 - count4_2;
            }
            DataTable table4 = mdbABC.Find("select * from 易损件信息 where 易损件所在工位='" + dateshuzu[1] + "'and 机台名称='"
                + dateshuzu[2] + "'and 易损件所在位置='" + dateshuzu[3] + "'and 易损件名称='" + dateshuzu[4] + "'");
            if (table4.Rows.Count > 0)
            {
                string sql1 = "update [易损件信息] set [易损件已使用次数]='" + count4_2 + "'," +
                    "[易损件剩余使用次数] = '" + count4_3 + "'" +
                    " where [机台名称] = '" + dateshuzu[2] + "'and 易损件名称='" + dateshuzu[4] + "'";
                var result1 = mdbABC.Change(sql1);
                if (result1 == true)
                {
                    this.BeginInvoke(ShowMsgAction, "工位：" + dateshuzu[1] + "机台：" + dateshuzu[2] +
                        "易损件信息更新：" + dateshuzu[3] + dateshuzu[4]);
                }
            }
            else
            {
                string sql1 = "insert into 易损件信息 (易损件所在工位, 机台名称, 易损件所在位置, 易损件名称, " +
                    "易损件理论使用次数,易损件已使用次数,易损件剩余使用次数)" +
                    " values ('" + dateshuzu[1] + "','" + dateshuzu[2] + "','" + dateshuzu[3] + "','" + dateshuzu[4] + "','"
                    + count4_1 + "','" + count4_2 + "','" + count4_3 + "')";
                bool result1 = mdbABC.Add(sql1.ToString());
                if (result1 == true)
                {
                    this.BeginInvoke(ShowMsgAction, "工位：" + dateshuzu[1] + "机台：" + dateshuzu[2] +
                        "新增易损件信息:" + dateshuzu[3] + dateshuzu[4]);
                }
            }
        }
        /// <summary>
        /// 统计信息ModPros
        /// </summary>
        /// <param name="dateshuzu"></param>
        List<Dictionary<string, List<string>>> listDic = new List<Dictionary<string, List<string>>>();
       
        //Dictionary<string, List<string>> dicmap = null;
        private void NewMethod3(string[] dateshuzu)
        {
            if (textBox8.Text.Trim().Length == 0)
            {
                this.BeginInvoke(ShowMsgAction, "输入最后一个机台！！！");
                return;
            }
            if (listDic.Count == 0)
            {
                Dictionary<string, List<string>> dicmapA = new Dictionary<string, List<string>>();
                dicmapA.Add(dateshuzu[1], ListDicodmh(dateshuzu));
                listDic.Add(dicmapA);
                /*if (dicmapA.Count == productTable.Rows.Count)
                {
                    double count1 = productTable.Rows.Count;//工位数量
                    double num1 = 0;//工序
                    double num2 = 0;//工序总和
                    double num3 = 0;//瓶颈;
                    double num4 = 0;//直通率
                    double num5 = 1; //直通率乘积
                    double timenum1 = 0;//利用时间
                                        //  double timenum2 = 0;//利用时间总和
                    double timenum3 = 0;//负荷时间
                                        // double timenum4 = 0;//负荷时间总和
                    foreach (List<string> valuelist in dicmapA.Values)
                    {
                        //0工单号//1成品名称//2工单数量
                        //3完成数量//4完成率//5合格率//6整体节拍//7工序时间//8利用时间 //9负荷时间
                        //10生产产品数量（总数）//11直通率
                        num2 += num1;
                        double.TryParse(valuelist[7], out num3);
                        if (num3 > num1)//计算最小的
                        {
                            double.TryParse(valuelist[7], out num1);//工序
                            num3 = num1;
                        }
                        double.TryParse(valuelist[11], out num4);
                        if (num4 != 0)//直通率乘积
                        {
                            num5 *= num4;
                        }
                        double.TryParse(valuelist[8], out timenum1);
                        double.TryParse(valuelist[9], out timenum3);
                    }
                    DataRowView dataRowView = (DataRowView)comboBox2.SelectedItem;
                    if (dicmapA.ContainsKey(dataRowView["名称"].ToString()))
                    {
                        List<string> dicstrkey = dicmapA[dataRowView["名称"].ToString()];
                        string sql3_1 = "insert into 统计信息 (工单号, 成品名称, 工单数量, 完成数量, 完成率, 合格率,整线节拍" +
          ",线平衡,OEE,直通率,更新时间,更新标识 )  values ('";
                        //0工单号//1成品名称//2工单数量
                        //3完成数量//4完成率//5合格率//6整体节拍//7工序时间//8利用时间 //9负荷时间
                        //10生产产品数量（总数）//11直通率
                        sql3_1 += dicstrkey[0] + "','";
                        sql3_1 += dicstrkey[1] + "','";
                        sql3_1 += dicstrkey[2] + "','";
                        sql3_1 += dicstrkey[3] + "','";
                        sql3_1 += dicstrkey[4] + "','";
                        sql3_1 += dicstrkey[5] + "','";
                        sql3_1 += dicstrkey[6] + "','";
                        if (num2 != 0 && num3 != 0 && count1 != 0)
                        {
                            sql3_1 += num2 / (count1 / num3) + "','";
                        }
                        else
                        {
                            sql3_1 += "  " + "','";
                        }
                        string OEE = " ";
                        double.TryParse(dicstrkey[8], out timenum1);//利用时间
                        double.TryParse(dicstrkey[9], out timenum3);//负荷时间
                        double rowcount9 = 0;//生产产品数量（总数）
                        double.TryParse(dicstrkey[10], out rowcount9);
                        double textcount9 = 0;//设计数度
                        double.TryParse(textBox9.Text, out textcount9);
                        double numhg = 0;//合格率
                        double.TryParse(dicstrkey[5], out numhg);
                        if (timenum1 != 0 && timenum3 != 0 && textcount9 != 0 && textcount9 != 0 && numhg != 0)
                        {
                            OEE = (timenum1 / timenum3) * (rowcount9 / (timenum1 * textcount9)) * numhg + "";
                        }
                        sql3_1 += OEE + "','";
                        sql3_1 += num5 + "','";
                        sql3_1 += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','";
                        sql3_1 += "F" + "')";
                        var result3_1 = mdb.Add(sql3_1.ToString());
                        if (result3_1)
                        {
                            this.BeginInvoke(ShowMsgAction, "添加一条统计信息");
                            listDic.Remove(dicmapA);
                        }
                    }
                }*/
            }
            else
            {
                for (int i = 0; i < listDic.Count; i++)
                {
                    Dictionary<string, List<string>> dicmap = new Dictionary<string, List<string>>();
                    dicmap = listDic[i];
                    if (dicmap.Count == productTable.Rows.Count)
                    {
                        double count1 = productTable.Rows.Count;//工位数量
                        double num1 = 0;//工序
                        double num2 = 0;//工序总和
                        double num3 = 0;//瓶颈;
                        double num4 = 0;//直通率
                        double num5 = 1; //直通率乘积
                        double timenum1 = 0;//利用时间
                                            //  double timenum2 = 0;//利用时间总和
                        double timenum3 = 0;//负荷时间
                                            // double timenum4 = 0;//负荷时间总和
                        foreach (List<string> valuelist in dicmap.Values)
                        {
                            //0工单号//1成品名称//2工单数量
                            //3完成数量//4完成率//5合格率//6整体节拍//7工序时间//8利用时间 //9负荷时间
                            //10生产产品数量（总数）//11直通率
                            num2 += num1;
                            double.TryParse(valuelist[7], out num3);
                            if (num3 > num1)//计算最小的
                            {
                                double.TryParse(valuelist[7], out num1);//工序
                                num3 = num1;
                            }
                            double.TryParse(valuelist[11], out num4);
                            if (num4 != 0)//直通率乘积
                            {
                                num5 *= num4;
                            }
                            double.TryParse(valuelist[8], out timenum1);
                            double.TryParse(valuelist[9], out timenum3);
                        }
                       // DataRowView dataRowView = (DataRowView)comboBox2.SelectedItem;
                        if (dicmap.ContainsKey(textBox8.Text))
                        {
                            List<string> dicstrkey = dicmap[textBox8.Text.ToString()];
                            string sql3_1 = "insert into 统计信息 (工单号, 成品名称, 工单数量, 完成数量, 完成率, 合格率,整线节拍" +
              ",线平衡,OEE,直通率,更新时间,更新标识 )  values ('";
                            //0工单号//1成品名称//2工单数量
                            //3完成数量//4完成率//5合格率//6整体节拍//7工序时间//8利用时间 //9负荷时间
                            //10生产产品数量（总数）//11直通率
                            sql3_1 += dicstrkey[0] + "','";
                            sql3_1 += dicstrkey[1] + "','";
                            sql3_1 += dicstrkey[2] + "','";
                            sql3_1 += dicstrkey[3] + "','";
                            sql3_1 += dicstrkey[4] + "','";
                            sql3_1 += dicstrkey[5] + "','";
                            sql3_1 += dicstrkey[6] + "','";
                            if (num2 != 0 && num3 != 0 && count1 != 0)
                            {
                                double AA = num2 / (count1 / num3);
                                sql3_1 += AA + "','";
                            }
                            else
                            {
                                sql3_1 += "0" + "','";
                            }
                            string OEE = "";
                            double.TryParse(dicstrkey[8], out timenum1);//利用时间
                            double.TryParse(dicstrkey[9], out timenum3);//负荷时间
                            double rowcount9 = 0;//生产产品数量（总数）
                            double.TryParse(dicstrkey[10], out rowcount9);
                            double textcount9 = 0;//设计数度
                            double.TryParse(textBox9.Text, out textcount9);
                            double numhg = 0;//合格率
                            double.TryParse(dicstrkey[5], out numhg);
                            if (timenum1 != 0 && timenum3 != 0 && textcount9 != 0 && textcount9 != 0 && numhg != 0)
                            {
                                double OEE1 = ((timenum1 / timenum3) * (rowcount9 / (timenum1 * textcount9)) * numhg);
                                OEE = OEE1.ToString();
                            }
                            else
                            {
                                OEE = "0";
                            }
                            sql3_1 += OEE + "','";
                            sql3_1 += num5 + "','";
                            sql3_1 += DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','";
                            sql3_1 += "F" + "')";
                            var result3_1 = mdbABC.Add(sql3_1.ToString());
                            if (result3_1)
                            {
                                this.BeginInvoke(ShowMsgAction, "添加一条统计信息");
                                listDic.Remove(dicmap);
                            }
                        }
                    }
                    else if (dicmap.ContainsKey(dateshuzu[1]))
                    {
                        Dictionary<string, List<string>> dicmapA = new Dictionary<string, List<string>>();
                        dicmapA.Add(dateshuzu[1], ListDicodmh(dateshuzu));
                        listDic.Add(dicmapA);
                        break;
                    }
                    else
                    {
                        dicmap.Add(dateshuzu[1], ListDicodmh(dateshuzu));
                        listDic[i] = dicmap;
                        break;
                    }
                }
            }
            
        }

        private List<string> ListDicodmh(string[] dateshuzu)
        {
            //3+机台名称1+工单号2+工单数量3+完成数量4+
            //完成率5+合格率6+整体节拍7+生产产品数量（总数）8
            //+ 工序时间9+利用时间10+负荷时间11+直通率12+成品名称13
            List<string> list = new List<string>();
            list.Add(dateshuzu[2]);//工单号
            list.Add(dateshuzu[13]);//成品名称
            list.Add(dateshuzu[3]);//工单数量
            list.Add(dateshuzu[4]);//完成数量
            list.Add(dateshuzu[5]);//完成率
            list.Add(dateshuzu[6]);//合格率
            list.Add(dateshuzu[7]);//整体节拍
            list.Add(dateshuzu[9]);//工序时间
            list.Add(dateshuzu[10]);//利用时间
            list.Add(dateshuzu[11]);//负荷时间
            list.Add(dateshuzu[8]);//生产产品数量（总数）
            list.Add(dateshuzu[12]);//直通率
            return list;
        }

        /// <summary>
        /// 生成信息表
        /// </summary>
        /// <param name="dateshuzu"></param>
        /// <param name="conn"></param>
        private void NewMethod2(string[] dateshuzu)
        {
            /// 2+工位名称+当前工单号+产品条码+操作人员+则式时间+则试结果+ 则试节拍+则试项名称+ 则试项上限+则试项下限+测式项实际值
            string sql2 = "insert into 生产信息 (工位名称 ,当前工单号, 产品条码,操作人员, " +
                "测试时间,测试结果,测试节拍,测试项名称," +
                "测试项上限,测试项下限,测试项的实际值, 更新标识)" +
                   " values ('" + dateshuzu[1] + "','" + dateshuzu[2] + "','" + dateshuzu[3] + "','" + dateshuzu[4] + "','"
               + dateshuzu[5] + "','" + dateshuzu[6] + "','" + dateshuzu[7] + "','" + dateshuzu[8] + "','"
               + dateshuzu[9] + "','" + dateshuzu[10] + "','" + dateshuzu[11] + "','" + "F" + "')";
            var result2 = mdbABC.Add(sql2.ToString());
        }
        /// <summary>
        /// 故障信息表
        /// </summary>
        /// <param name="dateshuzu"></param>
        private void NewMethod1(string[] dateshuzu)
        {
            /// 1+故障所在工位+机台名称+ 故障状态+故障的描述+触发故障的开始时间 
            /// 1+故障所在工位+机台名称+故障的描述+触发故障的结束时间
            DataTable table1 = mdbABC.Find("select * from 故障信息 where 故障发生工位='" + dateshuzu[1] + "'and 机台名称='"
                + dateshuzu[2] + "'and 故障描述='" + dateshuzu[4] + "'" + "and 发生时间='" + dateshuzu[5] + "'");
            if (table1.Rows.Count > 0)
            {
                //string sql = "update [Codes] set [Name]='" + funmae + "'" + " where [ID] = '" + pid + "'";
                string sql1 = "update [故障信息] set [结束时间]='" + dateshuzu[6] + "'" + " where 故障发生工位='" + dateshuzu[1] + "'and 机台名称='"
                + dateshuzu[2] + "'and 故障描述='" + dateshuzu[4] + "'" + "and 发生时间='" + dateshuzu[5] + "'";
                var result1 = mdbABC.Change(sql1);
                if (result1 == true)
                {
                    this.BeginInvoke(ShowMsgAction, "工位：" + dateshuzu[1] + "机台：" + dateshuzu[2] +
                        "故障停止信息：" + dateshuzu[3] + dateshuzu[6]);
                }
            }
            else
            {
                string sql1 = "insert into 故障信息 (故障发生工位, 机台名称, 故障类型, 故障描述, 发生时间, 结束时间, 更新标识)" +
                    " values ('" + dateshuzu[1] + "','" + dateshuzu[2] + "','" + dateshuzu[3] + "','" + dateshuzu[4] + "','"
                    + dateshuzu[5] + "','" + "   " + "','" + "F" + "')";
                bool result1 = mdbABC.Add(sql1.ToString());
                if (result1 == true)
                {
                    this.BeginInvoke(ShowMsgAction, "工位：" + dateshuzu[1] + "机台：" + dateshuzu[2] +
                        "故障发生信息：" + dateshuzu[3] + dateshuzu[4] + dateshuzu[5]);
                }
            }
        }

        #endregion
        /// <summary>
        /// 选择本地文件存放路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.label10.Text = path.SelectedPath;
        }
        /// <summary>
        /// 启动服务器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //// this.tbx_IP.Text = "127.0.0.2";
            //try
            //{
            //    Socket socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //    IPAddress ip = IPAddress.Parse(this.tbx_IP.Text);
            //    //创建对象端口
            //    IPEndPoint point = new IPEndPoint(ip, Convert.ToInt32(tbx_port.Text));
            //    this.tbx_IP.Text = ip.ToString();
            //    socketWatch.Bind(point);//绑定端口号
            //    ShowMsg("信息:监听成功!");
            //    socketWatch.Listen(100);//允许连接的客户端数量
            //                            //创建监听线程
            //    Thread thread = new Thread(Listen);
            //    thread.IsBackground = true;
            //    thread.Start(socketWatch);
            //    IsStart = true;
                

            //    // ShowBtnState();
            //}
            //catch (Exception ex)
            //{
            //    ShowMsg("错误信息:"+ex);
            //}
            
        }
        Dictionary<string, System.Net.Sockets.Socket> clientList = new Dictionary<string, System.Net.Sockets.Socket>();
        public void ConnSocket()
        {
            try
            {
                Socket socketWatch = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IPAddress ip = IPAddress.Parse(this.tbx_IP.Text);
                //创建对象端口
                IPEndPoint point = new IPEndPoint(ip, Convert.ToInt32(tbx_port.Text));
                this.tbx_IP.Text = ip.ToString();
                socketWatch.Bind(point);//绑定端口号
                ShowMsg("信息:监听成功!");
                socketWatch.Listen(100);//允许连接的客户端数量
                IsStart = true;

                Task.Factory.StartNew(() =>
                {
                    while (IsStart)
                    {
                        //接受接入的一个客户端
                        Socket connectClient = socketWatch.Accept();
                        if (connectClient != null)
                        {
                            string infor = connectClient.RemoteEndPoint.ToString();
                            clientList.Add(infor, connectClient);
                            //接收消息
                            ReciveMsg(connectClient);
                        }
                    }
                });

                // ShowBtnState();
            }
            catch (Exception ex)
            {
                ShowMsg("错误信息:" + ex);
            }

        }

        /// <summary>
        /// 接收消息
        /// </summary>
        /// <param name="client"></param>
        public void ReciveMsg(Socket client)
        {
            Task.Factory.StartNew(() =>
            {
                while (IsStart)
                {
                    try
                    {

                        ///定义服务器接收的字节大小
                        byte[] arrMsg = new byte[1024 * 1024 * 3];
                        ///接收到的信息大小(所占字节数)
                        int length = client.Receive(arrMsg);
                        Console.WriteLine("---" + length);
                        if (length > 0)
                        {
                            string recMsg = Encoding.UTF8.GetString(arrMsg, 0, length);
                            //获取客户端的端口号
                            IPEndPoint endPoint = client.RemoteEndPoint as IPEndPoint;
                            if (recMsg == "heartbeat")
                            {
                                ShowMsg("收到【" + client.RemoteEndPoint.ToString() + "】心跳：" + recMsg);
                                // 发送确认消息给客户端
                                Send(client, "OK");
                            }
                            else
                            {
                                //服务器显示客户端的端口号和消息
                                Task.Run(() =>
                                {
                                    Datapress(recMsg, client.RemoteEndPoint.ToString());
                                });
                            }
                        }

                    }
                    catch (Exception)
                    {
                        ///移除添加在字典中的服务器和客户端之间的线程
                        clientList.Remove(client.RemoteEndPoint.ToString());
                        ///关闭客户端
                        client.Close();
                        break;

                    }
                }
            });
        }


        /// <summary>
        /// 发送消息
        /// </summary>
        /// <param name="str"></param>
        void Send(Socket client, string str)
        {
            try
            {

                byte[] buffer = Encoding.UTF8.GetBytes(str);
                client.Send(buffer);
            }
            catch (Exception ex)
            {
                clientList.Remove(client.RemoteEndPoint.ToString());
                client.Close();
                ShowMsg(ex.Message);
                // MessageBox.Show("发送失败");
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            socketWatch?.Close();
            IsStart = false;
           // ShowBtnState();
            ShowMsg("信息:停止监听!");
        }
       /* private void btn_send_Click(object sender, EventArgs e)
        {
            if (!IsStart) return;
            string info = richTextBox1.Text;
            if (info == "") return;
            int index = richTextBox1.SelectedIndex;
            if (index < 0) return;
            if (ClientSockets.Count <= index) return;
            socketSend = ClientSockets[index];
            Send(info);
        }*/

        /// <summary>
        /// 等待接收客户端连接
        /// </summary>
        /// <param name="o"></param>
        void Listen(object o)
        {
            try
            {
                Socket socketWatch = o as Socket;
                while (IsStart)
                {
                    socketSend = socketWatch.Accept();//等待接收客户端连接
                    ClientSockets.Add(socketSend);
                    ClientIPPorts.Add(socketSend.RemoteEndPoint.ToString());
                    this.BeginInvoke(UpdateListViewDataAction);
                    this.BeginInvoke(ShowMsgAction, socketSend.RemoteEndPoint.ToString() + ":" + "连接成功!");
                    //开启一个新线程，执行接收消息方法
                    Thread r_thread = new Thread(Received);
                    r_thread.IsBackground = true;
                    r_thread.Start(socketSend);
                   // Thread t = new Thread(new ThreadStart(ScanOffline));
                  //  t.IsBackground = true;
                  //  t.Start();
                }
            }
            catch (Exception ex)
            {
                this.BeginInvoke(ShowMsgAction, "等待客户端监听发生异常:" + ex.Message);
            }
        }
        /// <summary>
        /// 扫描离线
        /// </summary>
        
        /// <summary>
        /// 服务器端不停的接收客户端发来的消息
        /// </summary>
        /// <param name="o"></param>
        void Received(object o)
        {
            try
            {
                Socket socketSend = o as Socket;
                if (!socketSend.Connected) return;
                while (IsStart)
                {
                    //客户端连接服务器成功后，服务器接收客户端发送的消息
                    byte[] buffer = new byte[1024 * 1024 * 3];
                    //实际接收到的有效字节数
                    int len = socketSend.Receive(buffer);
                    if (len == 0)
                    {
                        break;
                    }
                    string str = Encoding.UTF8.GetString(buffer, 0, len);
                    Task.Run(() =>
                    {
                        Datapress(str, socketSend.RemoteEndPoint.ToString());
                    });
                    
                    //SaveCSVlog(str);
                    this.BeginInvoke(ShowMsgAction, socketSend.RemoteEndPoint + ":" + str);

                }
            }
            catch (Exception ex)
            {
                
              //  mdbDatas mdbDa = new mdbDatas(path4);//conndnew connt
                string sql = "update  [product] set IP ='无' " +
                    ",conndnew='" + DateTime.Now + "'" +
                    ",connt='断开'  " +
                    "where IP='" + socketSend.RemoteEndPoint.ToString() + "'";
                foreach (DataRow row in productTable.Rows)
                {
                    // 名称, IP as IP,conndnew as 时间 ,connt as 状态
                    if (row["IP"].Equals(socketSend.RemoteEndPoint.ToString()))
                    {
                        //row["IP"] = SIP;
                        row["时间"] = DateTime.Now;
                        row["状态"] = "断开";
                    }
                }
                // bool resultC = mdbDa.Add(sql.ToString());
                // mdbDa.CloseConnection();
               // button8_Click(null, null);
                this.BeginInvoke(ShowMsgAction, socketSend.RemoteEndPoint+ "接收客户端内容发生异常:" + ex.Message);

            }
        }
        /// <summary>
        /// 发送消息
        /// </summary>
        /// <param name="str"></param>
        void Send(string str)
        {
            try
            {
                if (socketSend == null) return;
                foreach(Socket strip in ClientSockets)
                {
                    // Socket socket1=Ci
                    byte[] buffer = Encoding.UTF8.GetBytes(str);
                    strip.Send(buffer);
                }
                /*foreach(string key int new List<string>(ClientIPPorts)){ 
                }*/
                //if()
                //byte[] buffer = Encoding.UTF8.GetBytes(str);
               // socketSend.Send(buffer);
                MessageBox.Show("发送成功");
            }
            catch (Exception ex)
            {
                ShowMsg(ex.Message);
                MessageBox.Show("发送失败");
            }

        }
        private void ShowMsg(string msg)
        {
            Invoke(new Action(() =>
            {
                if (richTextBox1.TextLength > 50000)
                {
                    richTextBox1.Clear();
                }
                string info = string.Format("{0}:{1}\r\n", DateTime.Now.ToString("G"), msg);
                richTextBox1.AppendText(info);
            }));
        }
        private void ShowBtnState()
        {
            //btn_start.Enabled = !IsStart;
            //button7.Enabled = IsStart;
        }
        /// <summary>
        /// 用于实时监测客户端是否断开连接
        /// </summary>

        private void timer1_Tick(object sender, EventArgs e)
        {

            if (ClientSockets.Count == 0) return;
            for (int i = ClientSockets.Count - 1; i >= 0; i--)
            {
                if (!ClientSockets[i].Connected)
                {
                    ClientSockets.RemoveAt(i);
                    ClientIPPorts.RemoveAt(i);
                    this.BeginInvoke(UpdateListViewDataAction);
                }
            }
        }
        private void UpdateListViewData()
        {
           // listBox1.DataSource = null;
            //listBox1.DataSource = ClientIPPorts;
            //listBox1.Items.Clear();
            //for (int i = 0; i < ClientIPPorts.Count; i++)
            //{

            //}
        }

        public void SaveCSVlog(string log)
        {
            try
            {
                //  myJobManager.Run();
                if (System.IO.Directory.Exists("D:\\Logs") == false)
                {
                    System.IO.Directory.CreateDirectory("D:\\Logs");
                }
                // StringBuilder i = new StringBuilder();
                StringBuilder DataLine = new StringBuilder();

                string strT = DateTime.Now.Hour.ToString() + "时" + DateTime.Now.Minute.ToString() + "分" + DateTime.Now.Second.ToString() + "秒";

                //列标题
                // i.Append(log);
                //行数据
                DataLine.Append(strT + ":" + log);
                string FileName = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();
                string FilePath = "D:\\Logs" + "\\" + FileName + ".CSV";

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
        /// <summary>
        /// 刷新状态
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button8_Click(object sender, EventArgs e)
        {
            Invoke(new Action(() =>
            {
                mdb = new mdbDatas(path4);
                string strStat = "";
                //productTable = mdb.Find("select * from product");
                productTable = mdb.Find("select Mname as 名称, IP as IP,conndnew as 时间 ,connt as 状态 from product");
                dataGridView1.DataSource = productTable;
                //int i = 0;
               
                // dataGridView1.Rows[2].DefaultCellStyle.ForeColor = Color.Red;
                //  comboBox2.DataSource = tb;//displaymembe
                // comboBox2.DisplayMember = "名称";
                mdb.CloseConnection();
            }));
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Send("0+"+comboBox3.SelectedValue);
        }
        #region---------型号设置-----------
        //DataTable codesData;
        //Board  
        //ID WorkID BoardName BoardCode MaxBoardCode MinBoardCode BeatBoardCode ResultBoardCode
        private void SYS_BOARD()
        {
            button10_Click(null, null);
            DataGridViewButtonColumn butnCo = new DataGridViewButtonColumn();
            butnCo.HeaderText = "操作";
            butnCo.Text = "保存";
            butnCo.Name = "btnCol";
            butnCo.DefaultCellStyle.NullValue = "保存";
            //dataGridView2.Columns.Insert(0, buttonColumn);// dataGridView2.I
            dataGridView5.Columns.Add(butnCo);
            //table1.Columns.Add("操作");
            //buttonColumn.C += new DataGridViewCellEventHandler(buttonColumn_CellClick)
            //  Models();//查询名称
            //  button16_Click(null, null);//初始品名表、刷新
            // mdb.CloseConnection();
            // 再次创建一个新的列对象并设置其属删除
            DataGridViewButtonColumn anotrButCo = new DataGridViewButtonColumn();
            anotrButCo.HeaderText = "操作"; // 第二个按钮的标题文本
            anotrButCo.Name = "btnCol2"; // 第二个按钮的名称
            anotrButCo.DefaultCellStyle.NullValue = "删除";
            // buttonColumn.DefaultCellStyle.NullValue = "删除";
            // 将该列对象也添加到DataGridView控件中
            dataGridView5.Columns.Add(anotrButCo);
        }
        private void button10_Click(object sender, EventArgs e)
        {
            mdb = new mdbDatas(path4);
           
            DataTable codesData = mdb.Find("select ID as 编号 , Name as 型号 from Codes");
            dataGridView5.DataSource = codesData;
            comboBox3.DataSource = codesData;
            comboBox3.DisplayMember = "型号";
            comboBox3.ValueMember = "编号";
            // Codes  ID CName
            mdb.CloseConnection();
        }
       
        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Codes  ID CName
            if (e.RowIndex == -1)
            { return; }
            //删除
            if (dataGridView5.Columns[e.ColumnIndex].Name == "btnCol2")
            {
                //说明点击的列是DataGridViewButtonColumn列

                DataGridViewColumn column = dataGridView5.Columns[e.ColumnIndex];
                string pid = this.dataGridView5.Rows[e.RowIndex].Cells[2].Value.ToString();
                mdb = new mdbDatas(path4);
                bool bl = mdb.Del(" DELETE FROM  [Codes] WHERE [ID] = '" + pid + "'");
                if (bl == true)
                {
                    MessageBox.Show("删除成功");
                }
                mdb.CloseConnection();
                button10_Click(null, null);
                // MessageBox.Show("删除功"+column);
                // string username = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            }
            //保存
            if (dataGridView5.Columns[e.ColumnIndex].Name == "btnCol")
            {
                //说明点击的列是DataGridViewButtonColumn列
                DataGridViewColumn column = dataGridView5.Columns[e.ColumnIndex];
                string pid = this.dataGridView5.Rows[e.RowIndex].Cells[2].Value.ToString();
                string funmae = this.dataGridView5.Rows[e.RowIndex].Cells[3].Value.ToString();
                //int i = dateM.Rows.Count;
                if (string.IsNullOrWhiteSpace(pid))
                {
                    MessageBox.Show("编号不能为空！");
                    return;
                }
                mdb = new mdbDatas(path4);
                DataTable table1 = mdb.Find("select * from Codes where [ID] = '" + pid + "'");
                if (table1.Rows.Count > 0)
                {
                    string sql = "update [Codes] set [Name]='" + funmae + "'" + " where [ID] = '" + pid + "'";
                    var result = mdb.Change(sql);
                    if (result == true)
                    {
                        MessageBox.Show("修改成功");
                    }
                }
                else
                {
                    string sql = "insert into Codes ([ID],[Name]) values ('"
                        + pid + "','" + funmae + "')";
                    bool result = mdb.Add(sql.ToString());
                    if (result == true)
                    {
                        MessageBox.Show("新增成功");
                    }
                }

                mdb.CloseConnection();
                button10_Click(null, null);
                //MessageBox.Show("新增功" + pid);
                // string username = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            }
        }

        #endregion

        
    }
}
