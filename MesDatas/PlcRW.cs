using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HslCommunication;

namespace MesDatasCore
{
    public class PlcRW
    {
        static PlcRW()
        {
            if (!Authorization.SetAuthorizationCode("f562cc4c-4772-4b32-bdcd-f3e122c534e3"))
            {
                MessageBox.Show("授权失败！当前程序只能使用8小时！");
                return;
            }

        }

        public string IP { get; set; }
        public int Port { get; set; }

        public virtual bool Connect() { return true; }
        public virtual bool Close() { return false; }

        public virtual string ReadString(string address, ushort length) { return ""; }
        public virtual void WriteString(string address, string length) { }

    }
     
}
