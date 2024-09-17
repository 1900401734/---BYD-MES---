using MesDatasCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace BydMesTool
{
    public class MesTool : cToolBase
    {
        public MesTool() : base()
        {
            sName = "MesTool";
            ParameterForm = Parameter;
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

        FormBydMesTool Parameter = new FormBydMesTool();
         
        public override void ReceiveCommand(InteractiveEventArgs e)
        {
            if (e.InfoType == InfoType.Command)
            {
                string cmd = e.Value;
                if (e.Name == Parameter.MesUserCheck)
                {
                    //CheckMesUser
                    // out bool 验证结果, out string MES反馈, out string XMLOUT
                    bool 验证结果 = false;
                    string MES反馈 = "";
                    string XMLOUT = "";
                    Parameter.UsersVarify(out 验证结果, out MES反馈, out XMLOUT);
                    if (验证结果)
                    {
                        LogMes("用户验证OK");
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                    }
                    else
                    {
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                        LogMes("用户验证失败");
                    }

                }
                else if (e.Name == Parameter.MesBarCodeCheck)
                {
                    //CheckMesBarCode
                    // string 产品条码, out bool 验证结果, out string MES反馈, out string XMLOUT
                    //Parameter.BarCodeVarify();

                    string 产品条码 = ""; bool 验证结果 = false; string MES反馈 = ""; string XMLOUT = "";
                    //通过InteractiveEventArgs 进行参数传递


                    Parameter.BarCodeVarify(产品条码, out 验证结果, out MES反馈, out XMLOUT);
                    if (验证结果)
                    {
                        LogMes("条码验证OK");
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                    }
                    else
                    {
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                        LogMes("条码验证失败");
                    }

                }
                else if (e.Name == Parameter.MesBarCodeSend)
                {
                    // bool 测试结果, string 产品条码, string 文件版本, string 软件版本, string 测试项, out bool 验证结果, out string MES反馈, out string XMLOUT
                    //Parameter.UpDateToMes();
                    bool 测试结果 = false; string 产品条码 = ""; string 文件版本 = ""; string 软件版本 = ""; string 测试项 = ""; bool 验证结果 = false; string MES反馈 = ""; string XMLOUT = "";
                    //通过InteractiveEventArgs 进行参数传递


                    Parameter.UpDateToMes(测试结果, 产品条码, 文件版本, 软件版本, 测试项, out 验证结果, out MES反馈, out XMLOUT);
                    if (验证结果)
                    {
                        LogMes("条码上传OK");
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                    }
                    else
                    {
                        InteractiveEventArgs eOK = new InteractiveEventArgs();
                        SetEnque(eOK);
                        LogMes("条码上传失败");
                    }
                }
            }

        }

    }
     
}
