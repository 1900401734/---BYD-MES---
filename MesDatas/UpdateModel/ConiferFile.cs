using HslCommunication;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UltimateFileSoftwareUpdate.UpdateModel
{
    public class ConiferFile
    {
        public string IP { get; set; } // Name of the file

        public int Port { get; set; } // Name of the file

        public string Token { get; set; } // Name of the file

        public string Path { get; set; } // Name of the file

        public string PuthExe { get; set; } // Name of the file

        public string Version { get; set; } // Name of the file

        public static ConiferFile IniAddConfig()
        {
            ConiferFile conifer = new ConiferFile();
            conifer.IP = BDsqlNOTNULL(conifer.IP, "IP");
            string PortA =BDsqlNOTNULL(conifer.Port.ToString(), "Port");
            int PortB = 100;
            int.TryParse(PortA, out PortB);
            conifer.Port = PortB;
            conifer.Token = BDsqlNOTNULL(conifer.Token, "Token");
            conifer.Path = BDsqlNOTNULL(conifer.Path, "Path");
            conifer.Version = BDsqlNOTNULL(conifer.Version, "Version");
            return conifer;
        }

        public static string BDsqlNOTNULL(string s, string valid)
        {
            valid = GetConfigKey(valid);
            if (string.IsNullOrWhiteSpace(valid))
            {
                return s;
            }
            else
            {
                return valid;
            }
        }
        private static string AddconfigPath = Application.StartupPath + "\\UpdatDebug\\UpdateFile\\UpdateApp.config";
        public static string GetConfigKey(string key)
        {
            Configuration ConfigurationInstance = ConfigurationManager.OpenMappedExeConfiguration(new ExeConfigurationFileMap()
            {
                ExeConfigFilename = AddconfigPath
            }, ConfigurationUserLevel.None);

            if (ConfigurationInstance.AppSettings.Settings[key] != null)
                return ConfigurationInstance.AppSettings.Settings[key].Value;
            else
                return string.Empty;
        }
        private static string AddJson = Application.StartupPath + "\\UpdatDebug\\UpdateFile\\UpdateApp.json";
        public static ConiferFile GetJson()
        {
            try
            {
                string strjson=  File.ReadAllText(AddJson);
                //ConiferFile conifer= JsonSerializer.Deserialize<ConiferFile>(strJson); ;
                return JsonConvert.DeserializeObject<ConiferFile>(strjson);
            }catch (Exception ex)
            {
                return null;
            }
        }
        public static void SaveJson(ConiferFile conifer)
        {
            try
            {
                File.WriteAllText(AddJson, conifer.ToJsonString());
            }catch (Exception ex)
            {
                string msg = ex.ToString();
            }
        }
        public static int CompareVersions(string version1, string version2)
        {

            // 将版本号分割成数字数组  
            var versionParts1 = version1.Split('.').Select(int.Parse).ToArray();
            var versionParts2 = version2.Split('.').Select(int.Parse).ToArray();

            // 确定要比较的最大长度  
            int maxLength = Math.Max(versionParts1.Length, versionParts2.Length);

            // 逐一比较版本号的每一部分  
            for (int i = 0; i < maxLength; i++)
            {
                int part1 = i < versionParts1.Length ? versionParts1[i] : 0;
                int part2 = i < versionParts2.Length ? versionParts2[i] : 0;

                if (part1 < part2)
                    return -1;
                if (part1 > part2)
                    return 1;
            }

            // 如果所有部分都相等，则版本号相同  
            return 0;
        }
    }
}
