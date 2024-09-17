using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
    public class SytemInfoEntity
    {
        /// <summary>
        /// ip
        /// </summary>
        public string IP { get; set; }
        /// <summary>
        /// 端口
        /// </summary>
        public string Port { get; set; }
        /// <summary>
        /// 超时设置
        /// </summary>
        public string Timeout { get; set; }
        /// <summary>
        /// NcCode
        /// </summary>
        public string NcCode { get; set; }
        /// <summary>
        /// Opration
        /// </summary>
        public string Opration { get; set; }
        /// <summary>
        /// 密码
        /// </summary>
        public string Password { get; set; }
        /// <summary>
        /// Resource
        /// </summary>
        public string Resource { get; set; }
        /// <summary>
        /// Site
        /// </summary>
        public string Site { get; set; }
        /// <summary>
        /// url
        /// </summary>
        public string Url { get; set; }
        /// <summary>
        /// 用户名
        /// </summary>
        public string User { get; set; }
        /// <summary>
        /// 文件版本
        /// </summary>
        public string FileVersion { get; set; }
        /// <summary>
        /// 软件版本
        /// </summary>
        public string SoftwareVersion { get; set; }
        /// <summary>
        /// 用户验证
        /// </summary>
        public string UserCheckCmd { get; set; }
        /// <summary>
        /// 认证通过指令
        /// </summary>
        public string UserCheckPass { get; set; }
        /// <summary>
        /// 认证未通过指令
        /// </summary>
        public string UserCheckFail { get; set; }
        /// <summary>
        /// 条码响应指令
        /// </summary>
        public string CodeCheckCmd { get; set; }
        /// <summary>
        /// 条码验证通过指令
        /// </summary>
        public string CodeCheckPass { get; set; }
        /// <summary>
        /// 条码上传响应指令
        /// </summary>
        public string CodeSendCmd { get; set; }
    }
}
