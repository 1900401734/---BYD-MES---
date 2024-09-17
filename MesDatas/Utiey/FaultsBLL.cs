using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas.Utiey
{
    class FaultsBLL
    {
        // public List<Faults> faultsMap { get; set; }
        // 
        //"1+" + Namestation() + "+" + label54.Text + "+触发停机+"
        //+ faults + "+" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        public List<string> SETFaultsName(string stationName,string machiName, 
            string status,string readss, DataTable faultsTable)
        {
            List<string> faultsInform = new List<string>();
            Faults faults = new Faults();
            string[] faultsID = new string[] { };
            string faultsName = "";
            faultsID = readss.ToString().Split(new char[] { ',' });
            if (faultsID.Length > 0) {
                for (int i = 0; i < faultsID.Length; i++) {
                    faultsName= faults.SETFaultsName(faultsID[i], faultsTable);
                    //faults.stationName = stationName;
                    faultsInform.Add("1 + " + stationName + " + " + machiName + " + "+ status + "+ "
        + faultsName + "+");
                }
            }
            return faultsInform;
        }
       /* public Faults FaultsBalist(string readss,string stationName)
        {
            Faults faults = new Faults();
            string[] faultsID = new string[] { };
            faultsID= readss.ToString().Split(new char[] { ',' });
            faults.SETFaultsName(faultsID);
            faults.stationName = stationName;
            return faults;
        }*/
    }
}
