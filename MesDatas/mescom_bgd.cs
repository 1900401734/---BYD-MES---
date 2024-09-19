using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;
namespace 工艺部信息化组
{

    public struct CONFIG
    {
        public static string IP;
        public static string PORT;
        public static string URL;
        public static string Site;
        public static string UserName;
        public static string Password;
        public static string Resource;
        public static string Operation;
        public static string NcCode;
        public static int TimeOut;
    }
    public class BydMesCom
    {
        private static string ParamOUT;

        public static void 绑定工单(string 工单号, out bool 绑定结果, out string MES反馈, out string XMLOUT)
        {
            MES反馈 = BydMesCom.GetHtmlByPost("http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL, "&message=" + ("<PRODUCTION_REQUEST><RESOURCE_BANDING_SHOPORDER><SITE>" + CONFIG.Site + "</SITE><NAME>" + CONFIG.UserName + "</NAME><PWD>" + CONFIG.Password + "</PWD><RESOURCE>" + CONFIG.Resource + "</RESOURCE><SHOPORDER>" + 工单号 + "</SHOPORDER></RESOURCE_BANDING_SHOPORDER></PRODUCTION_REQUEST>"), CONFIG.TimeOut);
            绑定结果 = CutResult(MES反馈);
            XMLOUT = ParamOUT;
        }

        public static void 用户验证(out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
            MES反馈 = BydMesCom.GetHtmlByPost("http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL, "&message=" + ("<PRODUCTION_REQUEST><USER><SITE>" + CONFIG.Site + "</SITE><NAME>" + CONFIG.UserName + "</NAME><PWD>" + CONFIG.Password + "</PWD></USER></PRODUCTION_REQUEST>"), CONFIG.TimeOut);
            验证结果 = CutResult(MES反馈);
            XMLOUT = ParamOUT;
        }

        public static void 条码验证(string 产品条码, out bool 验证结果, out string MES反馈, out string XMLOUT)
        {
            MES反馈 = BydMesCom.GetHtmlByPost("http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL, "&message=" + ("<PRODUCTION_REQUEST><START><SFC_LIST><SFC><SITE>" + CONFIG.Site + "</SITE><ACTIVITY>XML</ACTIVITY><ID>" + 产品条码 + "</ID><RESOURCE>" + CONFIG.Resource + "</RESOURCE><OPERATION>" + CONFIG.Operation + "</OPERATION><USER>" + CONFIG.UserName + "</USER><QTY></QTY><DATE_TIME></DATE_TIME><COMPLEX>" + "N" + "</COMPLEX></SFC></SFC_LIST></START></PRODUCTION_REQUEST>!erpautogy03!1234567@byd"), CONFIG.TimeOut);
            Thread.Sleep(200);
            Application.DoEvents();
            验证结果 = CutResult(MES反馈);
            XMLOUT = ParamOUT;
        }

        public static void 条码上传(bool 测试结果, string 产品条码, string 文件版本, string 软件版本, string 测试项,
                                   out bool 验证结果, out string MES反馈, out string XMLOUT)
        {

            string mge = "";
            if (测试结果 == true)
            {
                mge = PassValidate(产品条码, 文件版本, 软件版本, 测试项);
            }
            else
            {
                mge = ErrorValidate(产品条码, 文件版本, 软件版本, 测试项);
            }
            MES反馈 = mge;
            验证结果 = CutResult(MES反馈);
            XMLOUT = ParamOUT;
        }

        private static string PassValidate(string 产品条码, string 文件版本, string 软件版本, string 测试项)
        {
            string url = "http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL;
            string Param = "&message=" + ("PASS<PRODUCTION_REQUEST><COMPLETE><SFC_LIST><SFC><SITE>" + CONFIG.Site + "</SITE><ACTIVITY>XML</ACTIVITY><ID>" + 产品条码 + "</ID><RESOURCE>" + CONFIG.Resource + "</RESOURCE><OPERATION>" + CONFIG.Operation + "</OPERATION><USER>" + CONFIG.UserName + "</USER><QTY>1</QTY><DATE_TIME></DATE_TIME><DATE_STARTED></DATE_STARTED></SFC></SFC_LIST></COMPLETE></PRODUCTION_REQUEST>!erpautogy03!1234567@byd!PASS," + 文件版本 + "," + 软件版本 + 测试项);
            // return BydMesCom.GetHtmlByPost("http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL, "&message=" + ("PASS<PRODUCTION_REQUEST><COMPLETE><SFC_LIST><SFC><SITE>" + CONFIG.Site + "</SITE><ACTIVITY>XML</ACTIVITY><ID>" + 产品条码 + "</ID><RESOURCE>" + CONFIG.Resource + "</RESOURCE><OPERATION>" + CONFIG.Operation + "</OPERATION><USER>" + CONFIG.UserName + "</USER><QTY>1</QTY><DATE_TIME></DATE_TIME><DATE_STARTED></DATE_STARTED></SFC></SFC_LIST></COMPLETE></PRODUCTION_REQUEST>!erpautogy03!1234567@byd!PASS," + 测试项), CONFIG.TimeOut);
            return BydMesCom.GetHtmlByPost(url, Param, CONFIG.TimeOut);
        }

        private static string ErrorValidate(string 产品条码, string 文件版本, string 软件版本, string 测试项)
        {
            string url = "http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL;
            string Param = "&message=" + ("ERROR<PRODUCTION_REQUEST><NC_LOG_COMPLETE><SITE>" + CONFIG.Site + "</SITE><OWNER TYPE=\"USER\">" + CONFIG.UserName + "</OWNER><NC_CONTEXT>" + 产品条码 + "</NC_CONTEXT><QTY></QTY><IDENTIFIER></IDENTIFIER><FAILURE_ID></FAILURE_ID><DEFECT_COUNT>1</DEFECT_COUNT><COMMENTS></COMMENTS><DATE_TIME></DATE_TIME><RESOURCE>" + CONFIG.Resource + "</RESOURCE><OPERATION>" + CONFIG.Operation + "</OPERATION><ROOT_CAUSE_OPER></ROOT_CAUSE_OPER><NC_CODE>" + CONFIG.NcCode + "</NC_CODE><ACTIVITY>XML</ACTIVITY></NC_LOG_COMPLETE></PRODUCTION_REQUEST>!erpautogy03!1234567@byd!ERROR," + 文件版本 + "," + 软件版本 + 测试项);
            //return BydMesCom.GetHtmlByPost("http://" + CONFIG.IP + ":" + CONFIG.PORT + CONFIG.URL, "&message=" + ("ERROR<PRODUCTION_REQUEST><NC_LOG_COMPLETE><SITE>" + CONFIG.Site + "</SITE><OWNER TYPE=\"USER\">" + CONFIG.UserName + "</OWNER><NC_CONTEXT>" + 产品条码 + "</NC_CONTEXT><QTY></QTY><IDENTIFIER></IDENTIFIER><FAILURE_ID></FAILURE_ID><DEFECT_COUNT>1</DEFECT_COUNT><COMMENTS></COMMENTS><DATE_TIME></DATE_TIME><RESOURCE>" + CONFIG.Resource + "</RESOURCE><OPERATION>" + CONFIG.Operation + "</OPERATION><ROOT_CAUSE_OPER></ROOT_CAUSE_OPER><NC_CODE>" + CONFIG.NcCode + "</NC_CODE><ACTIVITY>XML</ACTIVITY></NC_LOG_COMPLETE></PRODUCTION_REQUEST>!erpautogy03!1234567@byd!ERROR," + 测试项), CONFIG.TimeOut);
            return BydMesCom.GetHtmlByPost(url, Param, CONFIG.TimeOut);
        }

        private static bool CutResult(string html)  //Y return true : N return false
        {
            return html.Contains("</b>Y</td>") ? true : false;
            // return true;
        }

        private static string GetHtmlByPost(string URL, string Param, int TimeOut)
        {
            ParamOUT = Param;
            string str;
            try
            {
                byte[] bytes = Encoding.GetEncoding("GB2312").GetBytes(Param);
                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(URL);
                httpWebRequest.ContentType = "application/x-www-form-urlencoded";
                httpWebRequest.Accept = "*/*";
                httpWebRequest.UserAgent = "Mozilla/4.0(compatible;MSIE 6.0;Windows NT 5.1;SV1;Maxthon;.NET CLR 1.1.4322)";
                httpWebRequest.Method = "POST";
                httpWebRequest.ContentLength = (long)bytes.Length;
                httpWebRequest.Timeout = TimeOut;
                httpWebRequest.ServicePoint.Expect100Continue = false;
                Stream requestStream = httpWebRequest.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.GetEncoding("GB2312"));
                str = streamReader.ReadToEnd();
                streamReader.Close();
                httpWebRequest.Abort();
                httpWebResponse.Close();
            }
            catch (Exception ex)
            {
                str = ex.Message;
            }
            return str.Replace("&lt;", "<").Replace("&gt;", ">");
        }

    }
}

