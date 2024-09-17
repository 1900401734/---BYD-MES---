using MesDatasCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace PlcCommunication
{
    public class PlcCommunicationTool : cToolBase
    {
        public PlcCommunicationTool() : base()
        {
            // Parameter.button1.Click += SendOutMsg;
            sName = "PlcT";
            ParameterForm = Parameter;
            // string name = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name;
            ParameterForm.Text = "PlcTool";
            Parameter.SetTool(this);

        }


        public override void LoadParameters()
        {
            // throw new NotImplementedException();

            Parameter.LoadParameter();
        }

        public override void SaveParameters()
        {
            // throw new NotImplementedException();
            Parameter.SaveParameter();
        }

        FormPlcCommunication Parameter = new FormPlcCommunication();

        #region"plc read & write"

        PlcRW rwPlc = null;

        public void SetPlcType(string type)
        {
            if (rwPlc != null)
            {
                rwPlc.Close();
            }

            switch (type)
            {
                case "ModBus":
                    rwPlc = new cModbus();
                    break;
                case "FinsTcp":
                    rwPlc = new cFinsTcp();
                    break;
                default:
                    rwPlc = new cPlcTest();
                    break;

            }
        }

        public bool PlcConnect(string addr, string ip)
        {
            if (rwPlc != null)
            {
                rwPlc.IP = addr;
                rwPlc.Port = int.Parse(ip);
                return rwPlc.Connect();
            }
            return false;
        }
        public bool PlcClose()
        {
            if (rwPlc != null)
            {
                return rwPlc.Close();
            }
            return false;
        }

        public string PlcRead(string addr, int len)
        {
            if(rwPlc!=null)
            {
                return rwPlc.ReadString(addr, (ushort)len);
            }
            else
            {
                throw new Exception("rwPlc =null");
                //return null;
            }

        }

        public void PlcWrite(string addr, string value)
        {
            if(rwPlc!=null)
            {
                rwPlc.WriteString(addr, value);
            }
            else
            {
                throw new Exception("rwPlc =null");
                //return null;
            }
           
           
        }
        #endregion


        #region "actions"


        public int PlcRead_HandlerActionWithArgus(BaseActionArgus e)
        {

            PlcReadArgus plcReadArgus = e as PlcReadArgus;
            //read test data from plc
            string sValue = PlcRead(plcReadArgus.Addr, plcReadArgus.len);
            // update test to main frame
            InteractiveEventArgs e1 = new InteractiveEventArgs();
            e1.Result = true;
            e1.IsKey = false;
            e1.InfoType = InfoType.Content;
            e1.Name = e.ParameterName;
            e1.Value = sValue;

            SetEnque(e1);
            return 0;
        }

        public int PlcWrite_HandlerActionWithArgus(BaseActionArgus e)
        {

            PlcWriteArgus plcWriteArgus = e as PlcWriteArgus;
            //Write test data from plc
            PlcWrite(plcWriteArgus.Addr, plcWriteArgus.sValue);

            return 0;
        }

        /// <summary>
        /// 根据参数设置创建 PlcReadAction
        /// </summary>
        /// <param name="pd"></param>
        public void SetReadPlcActions(PlcData pd)
        {
            string sKey = pd.Name;
            PlcReadArgus e = new PlcReadArgus();

            e.Addr = pd.addr;


            if (!ToolActions.ContainsKey(sKey))
            {
                ActionBase plcRead = new ActionBase();

                plcRead.HandlerActionWithArgus += PlcRead_HandlerActionWithArgus;
                ToolActions.Add(sKey, plcRead);
            }
            ToolActions[sKey].SetArgus(e);
        }

        public void SetWritePlcActions(PlcData pd)
        {
            string sKey = pd.Name;
            PlcReadArgus e = new PlcReadArgus();
            if (!ToolActions.ContainsKey(sKey))
            {
                ActionBase plcRead = new ActionBase();

                plcRead.HandlerActionWithArgus += PlcWrite_HandlerActionWithArgus;
                ToolActions.Add(sKey, plcRead);
            }
            ToolActions[sKey].SetArgus(e);
        }
          
        #endregion
    }
    
    public class PlcReadArgus : BaseActionArgus
    {
        public string Addr;
        public int len;
    }

    public class PlcWriteArgus : BaseActionArgus
    {
        public string Addr;
        public string sValue;
    }

}
