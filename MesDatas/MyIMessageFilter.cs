using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MesDatas
{
    public class MyIMessageFilter : IMessageFilter
    {
        //0x0200:鼠标移动
        //0x0202:鼠标左键UP
        //0x201:鼠标左键Down
        //0x0203:鼠标左键双击
        //0x0205:鼠标右键UP
        //0x0204:鼠标右键Down
        //0x020a:鼠标滚轮
        //0x100:键盘按下
        public bool PreFilterMessage(ref Message m)
        {
            //throw new NotImplementedException();
            if (m.Msg == 0x0200 || m.Msg == 0x0202 || m.Msg == 0x0202 || m.Msg == 0x0203 || m.Msg == 0x0205 || m.Msg == 0x204 || m.Msg == 0x100)
            {
                if (Form1.timer != null)
                {
                    Console.WriteLine(DateTime.Now.ToString("G") + ":系统有动作");
                    Form1.iOperCount = 0;
                    Form1.timer.Stop();
                    Form1.timer.Start();
                    //System.Console.WriteLine("使用了鼠标和键盘");
                    //if (Form1.timer != null)
                    //{
                    //    Form1.timer.Stop();
                    //    Form1.timer.Start();
                    //    System.Console.WriteLine("定时器被重置");
                    //}
                }
            }
            return false;
        }
    }
}
