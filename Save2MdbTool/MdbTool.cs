using MesDatasCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Save2MdbTool
{
    /// <summary>
    /// 每月一个数据库，一个数据库一个表格
    /// </summary>
    public class MdbTool : cToolBase
    {
        mdbDatas mdb = null;

        public MdbTool() : base()
        {
            
            sName = "111";
            ParameterForm = Parameter;
            string name = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name;
            ParameterForm.Text = name;
            Parameter.SetTool(this);

        }

        private void Test()
        {
            mdbDatas.CreateAccessDatabase("D:\\tt.mdb");
            mdbDatas.CreateMDBTable("D:\\tt.mdb", "test040914", new System.Collections.ArrayList(new object[] { "ID", "Name", "tt", "66" }));
            mdb = new mdbDatas("D:\\tt.mdb");


            DataTable dt = new DataTable("test040914");
            DataColumn dc = new DataColumn();
            dt.Columns.Add(dc);
            dt.Columns.Add("ID1", typeof(string));
            DataColumn dcName = new DataColumn("Name", typeof(string));
            dt.Columns.Add(dcName);

            DataRow dr = dt.NewRow();
            dt.Rows.Add(dr);
            dt.Rows.Add();
            dt.Rows.Add("ssss", DateTime.Now);
            dt.Rows.Add(dt.Rows[0].ItemArray);
            dr[0] = "kkk";

            mdb.DatatableToMdb("test040914", dt);

            mdb.CloseConnection();
        }


        public override void LoadParameters()
        {
            // throw new NotImplementedException();
            Parameter.LoadParameter();
            LogMes("mdbtoo LoadParameters");
        }

        public override void SaveParameters()
        {
            // throw new NotImplementedException();
            Parameter.SaveParameter();
        }

        FormSave2MdbTool Parameter = new FormSave2MdbTool();




        /// <summary>
        /// 响应主框架发出的消息，并且发送指令给仪器
        /// </summary>
        /// <param name="e"></param>
        public override void ReceiveCommand(InteractiveEventArgs e)
        {
            if (e.InfoType == InfoType.Command)
            {
                SetEnque(e);
            }

        }


    }
}
