using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
    public class ConfigManager
    {
        private static string _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AccountManager");
        public static string fileName = "credentials.json";
        public static string filePath = Path.Combine(_configPath, fileName);
        private static JObject _config;

        private static void LoadConfig()
        {

            if (!File.Exists(filePath))
            {
                CreateJsonFile(_configPath, fileName, _config);
            }

            if (_config == null)
            {
                string json = File.ReadAllText(filePath);
                _config = JObject.Parse(json);
            }
        }

        public static string GetConfigValue(string key)
        {
            LoadConfig();
            return _config[key]?.ToString() ?? string.Empty;
        }

        public static void CreateJsonFile(string _configPath, string fileName, JObject content)
        {
            try
            {
                // 确保目录存在
                Directory.CreateDirectory(_configPath);

                // 初始化JSON内容
                JObject defaultConfig = new JObject()
                {
                    ["DevUsername"] = "dev",
                    ["DevPassword"] = "dev"
                };
                // 将 JSON 内容转换为格式化的字符串
                string jsonContent = JsonConvert.SerializeObject(defaultConfig, Formatting.Indented);

                // 写入文件
                File.WriteAllText(filePath, jsonContent);

                Console.WriteLine($"JSON file created successfully at: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
