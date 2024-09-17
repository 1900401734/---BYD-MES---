using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MesDatas
{
    /// <summary>
    /// 算法结果
    /// </summary>
    public class AlgResult
    {
        public string sfc
        {
            get { return tempsfc; }
            set { tempsfc = value; }
        }
        private string tempsfc = "02353SUL**B2008110000002558";

        public string station
        {
            get { return _tempStation; }
            set { _tempStation = value; }
        }
        private string _tempStation = string.Empty;

        public string machineNo
        {
            get { return tempmachineNo; }
            set { tempmachineNo = value; }
        }
        private string tempmachineNo = "EW077A";

        public string testDate
        {
            get { return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); }
        }

        public string testResult
        {
            get
            {
                if (IsFlatPass && IsLineStraitPass)
                    return "PASS";
                else
                    return "FAIL";
            }
        }

        public string equipData
        {
            get
            {
                string strRet = string.Empty;
                strRet += "Left_Straitness\\" + Left_Straitness.ToString() + ";";
                strRet += "Right_Straitness\\" + Right_Straitness.ToString() + ";";
                strRet += "PCV_Straightness\\" + PCV_Straightness.ToString() + ";";
                strRet += "FlatNess\\" + FlatNess.ToString() + ";";
                return strRet;
            }
        }


        [JsonIgnore]
        public double Left_Straitness { get; set; }

        [JsonIgnore]
        public double Right_Straitness { get; set; }

        [JsonIgnore]
        public double PCV_Straightness { get; set; }


        [JsonIgnore]
        public double FlatNess { get; set; }


        [JsonIgnore]
        public bool IsLineStraitPass { get; set; }

        [JsonIgnore]
        public bool IsFlatPass { get; set; }

    }


    public class MesJsonObject
    {
        public AlgResult[] equipList
        {
            get { return temp; }
            set { temp = value; }
        }
        private AlgResult[] temp = new AlgResult[1];
    }



    public class MesSfcRetData
    {
        public int code { get; set; }

        public string msg { get; set; }

        public string modelType { get; set; }
    }

    public class MesDockData
    {
        public int code { get; set; }

        public string msg { get; set; }
    }


    public class MesSystemInfo
    {
        private const string strCheckCom = "?command=checkSfcStatus&withNoPwd=1";
        private const string dockData = "?command=dockEquipData&withNoPwd=1";

        /// <summary>
        /// 数据查询
        /// </summary>
        /// <param name="url"></param>
        /// <param name="sfc"></param>
        /// <param name="station"></param>
        /// <returns></returns>
        public static MesSfcRetData CheckSfcAvailable(string url, string sfc, string station)
        {
            try
            {

                string tempsfc = "&param={\"sfc\":\"";
                string tempstation = "\"station\":\"";

                string strSend = tempsfc + sfc + "\"," + tempstation + station + "\"}";
                if (string.IsNullOrEmpty(url))
                    return null;
                url += strCheckCom;
                string strRet = HttpPostHelper.Post(url, strSend);
                if (string.IsNullOrEmpty(strRet))
                    return null;
                MesSfcRetData retData = JsonConvert.DeserializeObject<MesSfcRetData>(strRet);
                return retData ;
            }
            catch (Exception ex)
            {
               // Global.Log.Write(ex.ToString());
                return null;
            }
        }
        /// <summary>
        /// 发送测试结果数据
        /// </summary>
        /// <param name="url"></param>
        /// <param name="algResult"></param>
        /// <returns></returns>
        public static MesDockData SendTestResult(string url, AlgResult algResult)
        {
            if (string.IsNullOrEmpty(url) || null == algResult)
            {
                return null;
            }
            try
            {
                url += dockData;
                MesJsonObject item = new MesJsonObject();
                item.equipList[0] = algResult;
                string strResult = "&param=" + JsonConvert.SerializeObject(item);
                string strRet = HttpPostHelper.Post(url, strResult);
                MesDockData retData = JsonConvert.DeserializeObject<MesDockData>(strRet);
                return retData;

            }
            catch (Exception ex)
            {
                //Global.Log.Write(ex.ToString());
                return null;
            }
        }
    }


    /// <summary>
    /// 客户端格式
    /// </summary>
    public class HttpPostHelper
    {
        /// <summary>
        /// 发送信息到客户端
        /// </summary>
        /// <param name="url">连接的地址</param>
        /// <param name="postData">发送的内容</param>
        /// <returns></returns>
        public static string Post(string url, string postData)
        {
            return Post(url, postData, "application/x-www-form-urlencoded");
        }

        /// <summary>
        /// 发送信息到客户端
        /// </summary>
        /// <param name="url">连接的地址</param>
        /// <param name="postData">发送的内容</param>
        /// <param name="contentType">发送信息类型</param>
        /// <returns>返回信息</returns>
        public static string Post(string url, string postData, string contentType)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.ContentType = contentType;
                request.Method = "POST";
                request.Timeout = 300000;
                byte[] bytes = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = bytes.Length;
                Stream writer = request.GetRequestStream();
                writer.Write(bytes, 0, bytes.Length);
                writer.Close();
                object hh = request.GetResponse();
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                string result = reader.ReadToEnd();
                response.Close();
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString()); ;
                return string.Empty;
            }
        }

        /// <summary>
        /// 登录连接
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postData"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static string Post(string url, string postData, string userName, string password)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.ContentType = "text/html; charset=UTF-8";
            request.Method = "POST";

            string usernamePassword = userName + ":" + password;
            CredentialCache credentialCache =
                new CredentialCache { { new Uri(url), "Basic", new NetworkCredential(userName, password) } };
            request.Credentials = credentialCache;
            request.Headers.Add("Authorization",
                "Basic " + Convert.ToBase64String(new ASCIIEncoding().GetBytes(usernamePassword)));

            byte[] bytes = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = bytes.Length;
            Stream writer = request.GetRequestStream();
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.ASCII);
            string result = reader.ReadToEnd();
            response.Close();
            return result;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="url">请求的servlet地址，不带参数</param>
        /// <param name="postData"></param>
        /// <returns>请求的参数，key=value&key1=value1</returns>
        public static string doHttpPost(string url, string postData)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            SetHeaderValue(request.Headers, "Content-Type", "application/json");
            SetHeaderValue(request.Headers, "Accept", "application/json");
            SetHeaderValue(request.Headers, "Accept-Charset", "utf-8");
            request.Method = "POST";
            request.Timeout = 300000;

            byte[] bytes = Encoding.UTF8.GetBytes(postData);
            request.ContentLength = bytes.Length;
            Stream writer = request.GetRequestStream();
            writer.Write(bytes, 0, bytes.Length);
            writer.Close();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            string result = reader.ReadToEnd();
            response.Close();
            return result;
        }

        /// <summary>
        /// 偶发性超时时试看看
        /// </summary>
        /// <param name="url"></param>
        /// <param name="postData"></param>
        /// <returns></returns>
        public static string HttpPostForTimeOut(string url, string postData)
        {
            GC.Collect();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Timeout = 20 * 600 * 1000;


            ServicePointManager.Expect100Continue = false;
            ServicePointManager.DefaultConnectionLimit = 200;

            request.KeepAlive = false;
            request.ProtocolVersion = HttpVersion.Version10;

            Stream myRequestStream = request.GetRequestStream();
            StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.GetEncoding("utf-8")); //如果JSON有中文则是UTF-8
            myStreamWriter.Write(postData);
            myStreamWriter.Close(); //请求中止,是因为长度不够,还没写完就关闭了.

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string registerResult = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            return registerResult;
        }

        public static void SetHeaderValue(WebHeaderCollection header, string name, string value)
        {
            var property =
                typeof(WebHeaderCollection).GetProperty("InnerCollection",
                    BindingFlags.Instance | BindingFlags.NonPublic);
            if (property != null)
            {

                NameValueCollection collection = property.GetValue(header, null) as NameValueCollection;
                collection[name] = value;

            }
        }
    }

    public class Http_Post
    {
        public static string HttpPostSample(string ip,string Port,int TimeOut,string Url,string Xml_data,out string xmlOut)
        {
            xmlOut = Xml_data;
            string Return_str;
            try
            {
                byte[] bytes = Encoding.GetEncoding("GB2312").GetBytes(Xml_data);
                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://" + ip + ";" + Port + Url);
                httpWebRequest.ContentType = "application/x-www-form-urlencode";
                httpWebRequest.Method = "POST";
                httpWebRequest.ContentLength = (long)bytes.Length;
                httpWebRequest.Timeout = TimeOut;
                httpWebRequest.ServicePoint.Expect100Continue = false;
                Stream requestStream = httpWebRequest.GetRequestStream();
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream(), Encoding.GetEncoding("GB2312"));
                Return_str = streamReader.ReadToEnd();
                streamReader.Close();
                httpWebRequest.Abort();
                httpWebResponse.Close();
            }
            catch(Exception ex)
            {
                Return_str = ex.ToString();
            }
            return Return_str.Replace("&lt;", "<").Replace("&gt", ">");
        }
    }


}
