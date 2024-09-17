using NLog;
using NLog.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogConfig
{
    public class NLogHelperYY
    {
        /// <summary>
        /// NLog日志记录类
        /// </summary>
        public readonly static Logger logger = LogManager.GetCurrentClassLogger();
        /// <summary>
        /// NLog配置文件路径
        /// </summary>
        private readonly static string LogPath = AppDomain.CurrentDomain.BaseDirectory.ToString() + "/LogConfig/NLog.config";
        /// <summary>
        /// 初始化NLog配置文件
        /// </summary>
        public static void InitConfigNLog()
        {
            bool isExists = System.IO.File.Exists(LogPath);
            if (isExists)
            {
                LogManager.Configuration = new XmlLoggingConfiguration(LogPath);
            }
        }
    }
}
