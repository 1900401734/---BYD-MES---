using HslCommunication;
using HslCommunication.ModBus;
using HslCommunication.Profinet.Keyence;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MesDatas
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Cbx_工单_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Button14_Click(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        List<DeviceInfo> deviceInfos = new List<DeviceInfo>();
        //初始化
        private void Form2_Load(object sender, EventArgs e)
        {
            //初始化各机台信息
            
            DeviceInfo d1 = new DeviceInfo();
            d1.deviceName = "设备1";
            d1.deviceNo = "device1";
            d1.ip = "127.0.0.1";
            d1.prot = "501";
            d1.isCoon = false;
            deviceInfos.Add(d1);
            DeviceInfo d2 = new DeviceInfo();
            d2.deviceName = "设备2";
            d2.deviceNo = "device2";
            d2.ip = "127.0.0.1";
            d2.prot = "502";
            d2.isCoon = false;
            deviceInfos.Add(d2);
            DeviceInfo d3 = new DeviceInfo();
            d3.deviceName = "设备3";
            d3.deviceNo = "device3";
            d3.ip = "127.0.0.1";
            d3.prot = "503";
            d3.isCoon = false;
            deviceInfos.Add(d3);
        }
        Dictionary<string, DeviceInfo> dictThread = null;
        
        /// <summary>
        /// 开启监控
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            dictThread = new Dictionary<string, DeviceInfo>();
            //timer1_Tick(null, null);
            NewMethod1();
            //初始化定时器
            Timer timer = new Timer();
            //定时器间隔10秒
            timer.Interval = 10000;
            //定时器触发事件，timer1_Tick是winform自带的Timer控件的方法
            timer.Tick += timer1_Tick;
            //定时器可用
            timer.Enabled = true;
            //启动定时器
            timer.Start();
            
        }
        private KeyenceMcNet KeyenceMcNet;
        /// <summary>
        /// 定时器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void timer1_Tick(object sender, EventArgs e)
        {
            foreach (var str in deviceInfos)
            {
                try
                {
                    ModbusTcpNet busTcpCl = dictThread[str.deviceNo].modbusTcp;
                    OperateResult conn = busTcpCl.ConnectServer();
                    dictThread[str.deviceNo].modbusTcp = busTcpCl;//不管是否连接成功，都更新modbusTcp
                    if (conn.IsSuccess)
                    {
                        //dictThread[str.deviceNo].modbusTcp = busTcpCl;
                        dictThread[str.deviceNo].isCoon = true;
                        //dictThread.Add(str.deviceNo, deviceInfo);
                        // PatPZLhead(str, deviceInfo);
                        LogMsg("设备号：" + str.deviceNo + " IP：" + str.ip + " 端口:" + str.prot + "连接成功!");
                    }
                    else
                    {
                        //dictThread[str.deviceNo].modbusTcp = busTcpCl;
                        dictThread[str.deviceNo].isCoon = false;
                        //dictThread.Add(str.deviceNo, deviceInfo);
                        //PatPZLhead(str, deviceInfo);
                        LogMsg("设备号：" + str.deviceNo + " IP：" + str.ip + " 端口:" + str.prot + "连接失败!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            //1、先注册PLC连接
            // NewMethod1();
            //2、保存连接到键值对
        }
        /// <summary>
        /// 初始化Dictionary参数，并且注册PLC
        /// </summary>
        private void NewMethod1()
        {
            //1、先注册PLC连接
            foreach (var str in deviceInfos)
            {
                ModbusTcpNet busTcpClient = new ModbusTcpNet(str.ip, int.Parse(str.prot), 0x01);   // IP Address, Port Num;
                busTcpClient.AddressStartWithZero = true;
                busTcpClient.IsStringReverse = false;
                try
                {
                    DeviceInfo deviceInfo = new DeviceInfo();
                    deviceInfo.deviceName = str.deviceName;
                    deviceInfo.deviceNo = str.deviceNo;
                    deviceInfo.ip = str.ip;
                    deviceInfo.prot = str.prot;
                    OperateResult connect = busTcpClient.ConnectServer();
                    if (connect.IsSuccess)
                    {
                        deviceInfo.modbusTcp = busTcpClient;
                        deviceInfo.isCoon = true;
                        dictThread.Add(str.deviceNo, deviceInfo);
                        LogMsg("设备号：" + str.deviceNo + " IP：" + str.ip + " 端口:" + str.prot + "连接成功!");
                    }
                    else
                    {
                        deviceInfo.modbusTcp = busTcpClient;
                        deviceInfo.isCoon = false;
                        dictThread.Add(str.deviceNo, deviceInfo);
                        LogMsg("设备号：" + str.deviceNo + " IP：" + str.ip + " 端口:" + str.prot + "连接失败!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //2、保存连接到键值对
            }
        }

        private void NewMethod(DeviceInfo str, DeviceInfo deviceInfo)
        {
            if (dictThread.ContainsKey(deviceInfo.deviceNo))
            {
                dictThread[deviceInfo.deviceNo] = deviceInfo;
            }
            else
            {
                dictThread.Add(str.deviceNo, deviceInfo);
            }
        }

        /// <summary>
        /// 通知机台1#  AGV到位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values) 
            {
                if (str.deviceNo == "device1") 
                {
                    if (str.isCoon == true) 
                    {
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        busTcpClient.Write("1", 1);
                        LogMsg("设备号：" + str.deviceNo + " agv已到位!");
                    }
                }
            }
        }

        /// <summary>
        /// 写日志
        /// </summary>
        /// <param name="msg"></param>
        private void LogMsg(string msg)
        {
            this.Invoke(new Action(() =>
            {
                if (richTextBox1.TextLength > 50000)
                {
                    richTextBox1.Clear();
                }

                richTextBox1.AppendText(DateTime.Now.Hour.ToString() +
                    DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + "_" +
                    DateTime.Now.Millisecond.ToString() + ":" + msg + "\r\n");
                richTextBox1.ScrollToCaret();
                
            }));
        }
        /// <summary>
        ///  AGV到位 机台滚筒滚动
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button7_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values)
            {
                if (str.deviceNo == "device1")
                {
                    if (str.isCoon == true)
                    {
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        var Rea = busTcpClient.ReadInt32("1").Content;
                        if(Rea == 1)
                        {
                            //
                            busTcpClient.Write("2", 1);
                            LogMsg("机台1滚筒已经启动滚动");
                        }
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 机台1已经收到料
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button5_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values)
            {
                if (str.deviceNo == "device1")
                {
                    if (str.isCoon == true)
                    {
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        //if 
                        busTcpClient.Write("3", 1);
                        LogMsg(" 机台1已经收到，设备号：" + str.deviceNo + "料");
                    }
                }
            }
        }
        /// <summary>
        /// 通知机台1#  AGV2到位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button11_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values)
            {
                if (str.deviceNo == "device2")
                {
                    if (str.isCoon == true)
                    {  
                        //
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        busTcpClient.Write("1", 1);
                        LogMsg("设备号：" + str.deviceNo + " agv已到位!");
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values)
            {
                if (str.deviceNo == "device2")
                {
                    if (str.isCoon == true)
                    {
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        var Rea = busTcpClient.ReadInt32("1").Content;
                        if (Rea == 1)
                        {
                            //
                            busTcpClient.Write("2", 1);
                            LogMsg("机台1滚筒已经启动滚动");
                        }
                    }
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            foreach (DeviceInfo str in dictThread.Values)
            {
                if (str.deviceNo == "device2")
                {
                    if (str.isCoon == true)
                    {
                        ModbusTcpNet busTcpClient = str.modbusTcp;
                        //if AGV  
                        busTcpClient.Write("3", 1);
                        LogMsg(" 机台1已经收到，设备号：" + str.deviceNo + "料");
                    }
                }
            }
        }


    }

    public class DeviceInfo 
    {
        /// <summary>
        /// 设备名称
        /// </summary>
        public string deviceName { get; set; }
        /// <summary>
        /// 设备编号
        /// </summary>
        public string deviceNo { get; set; }
        /// <summary>
        /// IP
        /// </summary>
        public string ip { get; set; }
        /// <summary>
        /// 端口
        /// </summary>
        public string prot { get; set; }
        /// <summary>
        /// 是否链接
        /// </summary>
        public bool isCoon { get; set; }
        /// <summary>
        /// modubs对象
        /// </summary>
        public ModbusTcpNet modbusTcp { get; set; }
    }
    

}
