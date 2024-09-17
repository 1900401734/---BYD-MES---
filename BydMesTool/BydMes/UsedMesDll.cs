using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 工艺部信息化组;

namespace MesDatas.BydMes
{
    class UsedMesDll
    {
        public string Ip { get; set; }
        public string Port { get; set; }
        public string TimeOut { get; set; }
        public string Url { get; set; }
        public string Site { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Resource { get; set; }
        public string Operation { get; set; }
        public string NcCode { get; set; }
        private void Init()
        {
            工艺部信息化组.CONFIG.IP = Ip;
            工艺部信息化组.CONFIG.PORT = Port;
            工艺部信息化组.CONFIG.TimeOut = int.Parse(TimeOut);
            工艺部信息化组.CONFIG.URL = Url;
            工艺部信息化组.CONFIG.Site = Site;
            工艺部信息化组.CONFIG.UserName = UserName;
            工艺部信息化组.CONFIG.Password = Password;
            工艺部信息化组.CONFIG.Resource = Resource;
            工艺部信息化组.CONFIG.Operation = Operation;
            工艺部信息化组.CONFIG.NcCode = NcCode;
        }
       


        public bool UserVerify(out string MESren, out string XMLOUT)
        {
            string MESRETURN = "...";
            string XMLIN = "...";
            bool ok = false;
            BydMesCom.用户验证(out ok, out MESren, out XMLOUT);

            //MESRETURN = "接收的数据：\n" + MESren;
            //XMLIN = "发送的数据：\n" + XMLOUT;
            return ok;
        }

        public bool BarCodeVerify(out string MESren, out string XMLOUT)
        {
            bool ok = false;
            string MESRETURN = "...";
            string XMLIN = "...";
            string SFCIN = "";
            BydMesCom.条码验证(SFCIN, out ok, out MESren, out XMLOUT);

            //MESRETURN = "接收的数据：\n" + MESren;
            //XMLIN = "发送的数据：\n" + XMLOUT;
            return ok;
        }

        public void BarCodeSend(bool 测试结果, string 产品条码, string 文件版本, string 软件版本, string 测试项, out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
          
            BydMesCom.条码上传(测试结果, 产品条码, 文件版本, 软件版本, 测试项, out 验证结果, out MES反馈, out XMLOUT);

            //MESRETURN = "接收的数据：\n" + MESren;
            //XMLIN = "发送的数据：\n" + XMLOUT;
           // return MES反馈;
        }
    }
}
