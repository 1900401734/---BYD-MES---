using INIFile;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MesDatasCore
{
    public interface IMesDatasBase
    {
        /// <summary>
        /// 
        /// </summary>
        Dictionary<string, IToolBase> Tools { get; set; }
        /// <summary>
        /// 响应MesDatas的指令
        /// </summary>
        SendCommand sendCommand { get; set; }
    }
    public interface IToolBase
    {
        string sName { get; set; }
        string ToolDllName { get; }
        string ParameterRootPath { set; }
        IMesDatasBase MesDatasMain { set; }
        /// <summary>
        /// 收到数据后发给MesDatas
        /// </summary>
        ReceiveChanged receiveChanged { get; set; }
        /// <summary>
        /// 参数页面
        /// </summary>
        Form ParameterForm { get; }
        /// <summary>
        /// 参数加载
        /// </summary>
        void LoadParameters();
        /// <summary>
        /// 参数保存
        /// </summary>
        void SaveParameters();

        void ToolInit();


        string ReadString(string section, string key);
        bool WriteString(string section, string key, string value);
        bool IsRunLoop { get; set; }

        List<string> GetCommands();
    }

    /// <summary>
    /// 单个Tool
    /// </summary>
    public class cToolBase : IToolBase
    {
        public cToolBase()
        {
            ToolActions.Clear();
        }

        public void ToolInit()
        {
            Task.Run(new Action(() =>
            {
                LoopSendOut();
            }));
        }

        private string mName = "";
        public string sName { get { return mName; } set { mName = value; } }

        string IToolBase.ToolDllName { get { return "ToolTestDll"; } }

        private IMesDatasBase mMesDatasMain = null;

        public IMesDatasBase MesDatasMain
        {
            set
            {
                mMesDatasMain = value;
                if (mMesDatasMain != null)
                {
                    mMesDatasMain.sendCommand += ReceiveCommand;
                }
            }
        }

        public List<string> AllCommands()
        {
            List<string> cmds = new List<string>();
            foreach (KeyValuePair<string, IToolBase> tool in mMesDatasMain.Tools)
            {
                cmds.AddRange(tool.Value.GetCommands());
            }
            return cmds;
        }

        public List<string> GetCommands()
        {
            List<string> cmds = new List<string>();
            return cmds;
        }

        public ReceiveChanged receiveChanged
        {
            get
            {
                return ReceiveChangedSendOut;
            }
            set
            {
                ReceiveChangedSendOut = value;
            }
        }

        public Form ParameterForm { get; set; }

        public string ReadString(string section, string key)
        {
            return INIClass.INIGetStringValue(ParameterRootPath, section, key);
        }

        //object lockFileWrite = new object();
        public bool WriteString(string section, string key, string value)
        {
            bool b = false;
            // lock (lockFileWrite)
            {
                b = INIClass.INIWriteValue(ParameterRootPath, section, key, value);
            }


            return b;
        }



        public virtual void LoadParameters()
        {
            //throw new NotImplementedException();
        }

        public virtual void SaveParameters()
        {
            // throw new NotImplementedException();
        }

        public string ParameterRootPath
        {
            get { return mParameterRootPath; }
            set { mParameterRootPath = value; }
        }

        private string mParameterRootPath = "";

        private Queue<InteractiveEventArgs> eQueue = new Queue<InteractiveEventArgs>();

        private object lockObj = new object();

        public void SetEnque(InteractiveEventArgs e)
        {
            lock (lockObj)
            {
                eQueue.Enqueue(e);
            }
        }
        private bool mIsRunLoop = true;
        public bool IsRunLoop
        {
            get { return mIsRunLoop; }
            set { mIsRunLoop = value; }
        }
        /// <summary>
        /// 传出数据或者指令
        /// </summary>
        public void LoopSendOut()
        {
            while (IsRunLoop)
            {
                InteractiveEventArgs eArgs = null;
                if (eQueue.Count > 0)
                {
                    lock (lockObj)
                    {
                        eArgs = eQueue.Dequeue();
                    }
                    ReceiveChangedSendOut?.Invoke(eArgs);
                }
                Thread.Sleep(5);
                Application.DoEvents();
            }
        }

        public void LogMes(string Msg)
        {
            InteractiveEventArgs e = new InteractiveEventArgs();
            e.InfoType = InfoType.LogMsg;
            e.Value = Msg;
            SetEnque(e);
        }

        // FormParameter Parameter = new FormParameter();

        public ReceiveChanged ReceiveChangedSendOut;

        /// <summary>
        /// 要发送出去的消息
        /// </summary>
        public InteractiveEventArgs eArgs = new InteractiveEventArgs();

        public Dictionary<string, ActionBase> ToolActions = new Dictionary<string, ActionBase>();



        /// <summary>
        /// 响应主框架发出的消息，并且发送指令给仪器
        /// </summary>
        /// <param name="e"></param>
        public virtual void ReceiveCommand(InteractiveEventArgs e)
        {
            if (e.InfoType == InfoType.Command)
            {
                string cmd = e.Value;
                if (ToolActions.ContainsKey(cmd))
                {
                    // 发送数据给硬件-请求采集数据
                    // MessageBox.Show("收到指令：" + e.Value);
                    ToolActions[cmd].Run();
                }
            }

        }

    }




    public class InteractiveEventArgs : EventArgs
    {
        public string Name;
        public string Value;
        public bool Result;
        public InfoType InfoType;
        public bool IsKey;
        public string Param;
    }
    public enum InfoType
    {
        Command,
        Content,
        LogMsg
    }
    public delegate void ReceiveChanged(InteractiveEventArgs e);
    public delegate void SendCommand(InteractiveEventArgs e);

}
