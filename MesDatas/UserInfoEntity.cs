using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
    public class UserInfoEntity
    {
        /// <summary>
        /// 用户名
        /// </summary>
        public string Uuser { get; set; }
        /// <summary>
        /// 密码
        /// </summary>
        public string Upwd { get; set; }
        /// <summary>
        /// 权限类型
        /// </summary>
        public string Utype { get; set; }
        /// <summary>
        /// 登录类型
        /// </summary>
        public string loginType { get; set; }
        /// <summary>
        /// 姓名
        /// </summary>
        public string userName { get; set; }
        /// <summary>
        /// 卡id
        /// </summary>
        public string cardID { get; set; }
    }
    public class DataConverter
    {
        public static List<UserInfoEntity> ConvertDataTableToList(DataTable table)
        {
            List<UserInfoEntity> userInfoEntities = new List<UserInfoEntity>();

            foreach (DataRow row in table.Rows)
            {
                UserInfoEntity userInfoEntity = new UserInfoEntity
                {
                    Uuser = row["工号"].ToString(),
                    Upwd = row["用户密码"].ToString(),
                    Utype = row["用户权限"].ToString(),
                    //  loginType = row["厂牌UID"].ToString(), // 假设 "厂牌UID" 是登录类型
                    userName = row["用户名"].ToString(), // 假设 "工号" 是姓名
                    cardID = row["厂牌UID"].ToString() // 假设 "厂牌UID" 是卡id
                };

                userInfoEntities.Add(userInfoEntity);
            }

            return userInfoEntities;
        }
    }
}
