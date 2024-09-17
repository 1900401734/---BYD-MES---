using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MesDatas.SqlConverter
{

    public class SqlToJsonConverter
    {

        public static string ConvertToJSON(string sql)
        {
            sql = sql.Trim().ToUpper();

            if (sql.StartsWith("INSERT"))
            {
                return ConvertInsertToJSON(sql);
            }
            else if (sql.StartsWith("DELETE"))
            {
                return ConvertDeleteToJSON(sql);
            }
            else if (sql.StartsWith("UPDATE"))
            {
                return ConvertUpdateToJSON(sql);
            }

            throw new NotSupportedException("Unsupported SQL operation.");
        }

        private static string ConvertInsertToJSON(string sql)
        {
            // 示例: INSERT INTO TableName (Column1, Column2) VALUES ('Value1', 'Value2')
            var tableName = sql.Split(' ')[2];
            var columnsAndValues = sql.Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
            var columns = columnsAndValues[1].Split(',');
            var values = columnsAndValues[3].Replace("VALUES", "").Trim().Split(',');

            var dictionary = new Dictionary<string, object>();
            for (int i = 0; i < columns.Length; i++)
            {
                dictionary.Add(columns[i].Trim(), values[i].Trim().Trim('\''));
            }

            var result = new
            {
                Operation = "INSERT",
                // TableName = tableName,
                Data = dictionary
            };

            return JsonConvert.SerializeObject(result, Formatting.None);
        }

        private static string ConvertDeleteToJSON(string sql)
        {
            // 示例: DELETE FROM TableName WHERE Column1 = 'Value1'
            var tableName = sql.Split(' ')[2];
            var condition = sql.Substring(sql.IndexOf("WHERE") + 6);

            var result = new
            {
                Operation = "DELETE",
                TableName = tableName,
                Condition = condition
            };

            return JsonConvert.SerializeObject(result, Formatting.None);
        }

        private static string ConvertUpdateToJSON(string sql)
        {
            // 示例: UPDATE TableName SET Column1 = 'Value1', Column2 = 'Value2' WHERE Column3 = 'Value3'
            var tableName = sql.Split(' ')[1];
            var setPart = sql.Substring(sql.IndexOf("SET") + 4, sql.IndexOf("WHERE") - sql.IndexOf("SET") - 4).Trim();
            var condition = sql.Substring(sql.IndexOf("WHERE") + 6);

            var sets = setPart.Split(',');

            var dictionary = new Dictionary<string, object>();
            foreach (var set in sets)
            {
                var keyValue = set.Split('=');
                dictionary.Add(keyValue[0].Trim(), keyValue[1].Trim().Trim('\''));
            }

            var result = new
            {
                Operation = "UPDATE",
                TableName = tableName,
                Data = dictionary,
                Condition = condition
            };

            return JsonConvert.SerializeObject(result, Formatting.None);
        }

        public static string ConvertToCustomLog(string json)
        {
            JObject jsonObject = JObject.Parse(json);
            string operation = jsonObject["Operation"]?.ToString();

            if (operation == "INSERT")
            {
                return ConvertInsertToLog(jsonObject["Data"] as JObject);
            }

            else if (operation == "DELETE")
            {
                return ConvertDeleteToLog(jsonObject["Data"] as JObject);
            }

            else if (operation == "UPDATE")
            {
                return ConvertUpdateToLog(jsonObject["Data"] as JObject);
            }

            throw new NotSupportedException("Unsupported operation.");
        }

        private static string ConvertInsertToLog(JObject data)
        {
            // 忽略用户密码项
            data.Remove("[用户密码]");

            string logMessage = "【用户新增】\n";
            logMessage += $"姓名：{data["[用户名]"]} | ";
            logMessage += $"工号：{data["[工号]"]} | ";
            logMessage += $"权限：{data["[用户权限]"]} | ";
            logMessage += $"厂牌UID：{data["[厂牌UID]"]}";

            return logMessage;
        }

        private static string ConvertDeleteToLog(JObject data)
        {
            string logMessage = "【用户删除】\n";
            logMessage += $"姓名：{data["[用户名]"]} | ";
            logMessage += $"工号：{data["[工号]"]} | ";
            logMessage += $"权限：{data["[用户权限]"]} | ";
            logMessage += $"厂牌UID：{data["[厂牌UID]"]}";

            return logMessage;
        }

        private static string ConvertUpdateToLog(JObject data)
        {
            string oldPermission = data["[OldPermission]"]?.ToString();
            string newPermission = data["[用户权限]"]?.ToString();
            string userName = data["[用户名]"]?.ToString();

            string logMessage = "【权限修改】\n";
            logMessage += $"已将{userName}的权限变更为{newPermission}！";

            return logMessage;
        }
    }
}
