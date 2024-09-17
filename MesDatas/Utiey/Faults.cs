using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas.Utiey
{
    class Faults
    {
        //public string stationName { get; set; }
        public string faultsName { get; set; }

        public string SETFaultsName(string faultsID, DataTable faultsTable)
        {
            foreach (DataRow row in faultsTable.Rows)
            {
                if (row["编号"].Equals(faultsID))
                {
                    faultsName = row["故障描述"].ToString();
                    break;
                }
            }
            if (string.IsNullOrWhiteSpace(faultsName))
            {
                faultsName = "没有添加该故障描述！！！";
            }
            return faultsName;
        }

        //"1+" + Namestation() + "+" + label54.Text + "+触发停机+"
        //+ faults + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
       
    }
}
