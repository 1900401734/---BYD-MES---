using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MesDatasCore;

namespace ToolSample
{
    public class ToolTest : cToolBase
    {
        public ToolTest() : base()
        {
            Parameter.button1.Click += SendOutMsg;
            sName = "111";
            ParameterForm = Parameter;
            string name = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name;
            ParameterForm.Text = name;
            Parameter.SetTool(this);
        }


        public override void LoadParameters()
        {
            Parameter.LoadParameter();
        }
        

        public override void SaveParameters()
        {
            Parameter.SaveParameter();
        }         

        FormToolSample Parameter = new FormToolSample();

        /// <summary>
        /// 发送数据或者指令给主框架
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public   void SendOutMsg(object sender, EventArgs e)
        {
            eArgs.Value = "this info is from tool test";
            SetEnque(eArgs);
        }


        /// <summary>
        /// 响应主框架发出的消息，并且发送指令给仪器
        /// </summary>
        /// <param name="e"></param>
        //public override void ReceiveCommand(InteractiveEventArgs e)
        //{
        //    if (e.infoType == InfoType.Command)
        //    {
        //        if (e.Value == Command)
        //        {
        //            // 发送数据给硬件-请求采集数据
        //            MessageBox.Show("收到指令：" + e.Value);
        //        }
        //    }

        //}


    }
}
