using HslCommunication.Enthernet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YYUltimateFileClient
{
    public class UltimateFileClientFactory
    {
        public IntegrationFileClient integrationFileClient;

        public void UltimateFileClientStart(string ip="127.0.0.1",int port=34567,string token= "A8826745-84E1-4ED4-AE2E-D3D70A9725B5")
        {
           
            // 创建客户端对象
            //TODO: Implement the client factory
            integrationFileClient = new IntegrationFileClient()
            {
                ConnectTimeOut = 10000,
                ServerIpEndPoint = new System.Net.IPEndPoint(System.Net.IPAddress.Parse(ip), port),
                Token = new Guid(token)                                         // 指定一个令牌
            };
            // integrationFileClient.Close();
            // 创建本地文件存储的路径
            //string path = Application.StartupPath + @"\Files";
            //if (!System.IO.Directory.Exists(path))
            //{
            //    System.IO.Directory.CreateDirectory(path);
            //}
        }
    }
}
