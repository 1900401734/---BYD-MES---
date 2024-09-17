using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Management;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;
//using SocketConnTimeOut;
//using Helper.Socket;
using System.Collections.Concurrent;

  namespace MesDatas
{
    class cWriteResult
    {
        private void WriteResult(AlgResult result,string MesURL)
        {
            MesSfcRetData ret = MesSystemInfo.CheckSfcAvailable(MesURL, result.sfc, result.station);
            if (ret == null)
            {
               // Global.Log.Write(string.Format("MESURL{0}错误 获取sfc{1} 运行", Global.MesSaveInfo.MesURL, result.sfc));
                return;
            }
            else if (ret.code != 0)
            {
               // Global.Log.Write(string.Format("发送失败{0}", ret.msg));
                return;
            }
            MesDockData Dockret = MesSystemInfo.SendTestResult(MesURL, result);
            if (Dockret == null)
            {
               // Global.Log.Write(string.Format("MESURL{0}错误 获取sfc{1} 运行", Global.MesSaveInfo.MesURL, result.sfc));
                return;
            }
            else if (ret.code != 0)
            {
               // Global.Log.Write(string.Format("发送测试结果到mes{0}", ret.msg));
                return;
            }
           // Global.Log.Write(string.Format("发送测试结果到mes成功{0}", ret.msg));
        }

    }
}
 
  