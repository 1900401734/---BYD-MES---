using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
    [Serializable]
    public class MesSaveInfo
    {
        /// <summary>
        /// 是否开启Mes
        /// </summary>
        public bool IsSendMes
        {
            get { return issendMes; }
            set { issendMes = value; }
        }
        private bool issendMes = false;

        public string MesURL 
        {
            get { return mesUrl; }
            set { mesUrl = value; }
        }
        private string mesUrl = "http://10.6.78.115:80/shywsbq/BydTraceSfcEquipData.do";

        public string MachineStation
        {
            get { return strMachineStaion; }
            set { strMachineStaion = value; }
        }
        private string strMachineStaion = "RCV";// "PlatnessTest";

        public string MachineNo
        {
            get { return strMachine; }
            set { strMachine = value; }
        }
        private string strMachine = "hengsheng001";

        public void Save()
        {
            string strErr = string.Empty;
            XmlSerializerHelper.SaveUtf8NoNamespace(MesFile,this,ref strErr);
            if (!string.IsNullOrEmpty(strErr))
            {
                //Global.Log.Write(strErr);
            }
        }

        public static MesSaveInfo LoadFile()
        {
            if (!File.Exists(MesFile))
                return null;
            string strErro = string.Empty;
            MesSaveInfo mes = XmlSerializerHelper.Load<MesSaveInfo>(MesFile,ref strErro);
            return mes;
        }

        public static string MesFile = AppDomain.CurrentDomain.BaseDirectory + "MesInfo.xml";
    }
}
