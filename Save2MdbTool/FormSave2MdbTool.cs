using MesDatasCore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Save2MdbTool
{
    public partial class FormSave2MdbTool : Form
    {
       
        public FormSave2MdbTool()
        {
            InitializeComponent();

          
        }
        OleDbDataAdapter adp;
        /// <summary>
        /// 将DateGridView中的数据更新到数据库
        /// </summary>
        public void DateToAccess()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            //string str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\\0000000\\2021年11月生产数据.mdb;Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            //OleDbConnection conn = new OleDbConnection(str);
            adp = new OleDbDataAdapter("select * from "+ cbx_Model.Text, conn);
            DataTable dtc = new DataTable();
            adp.Fill(dtc);
            dtc.Clear();
            dataGridViewDynamic1.DataSource = dtc;
           

        }

        private cToolBase cTool = null;
        public void SetTool(cToolBase Tool)
        {
            cTool = Tool;
        }

        private string ReadString(string key)
        {
            return cTool.ReadString(cTool.sName, key);
        }

        private void WriteString(string key, string value)
        {
            cTool.WriteString(cTool.sName, key, value);
        }

        public void LoadParameter()
        {
            // ReadString
            textBoxPath.Text = ReadString("Path");
            textBoxCmd.Text = ReadString("Command");
        }


        public void SaveParameter()
        {
            // WriteString
            WriteString("Path", textBoxPath.Text);
            WriteString("Command", textBoxCmd.Text);
        }
        /// <summary>
        /// 查询
        /// </summary>
        public void chaxun()
        {
            DateTime times_Start =dateTimePicker1.Value.Date;
            DateTime times_End = dateTimePicker2.Value.Date;
            string times_Start_string =times_Start.ToString();
            string times_End_string = times_End.ToString();
            DateTime times_Start_Sub = DateTime.Parse(times_Start_string.Substring(0,10) + " 00:00:00");
            DateTime times_End_Sub = DateTime.Parse(times_End_string.Substring(0, 10) + " 23:59:59");
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            OleDbParameter[] pars = new OleDbParameter[] 
            {
                new OleDbParameter("@column1",textBox_Code.Text),

                new OleDbParameter("@StartDate",times_Start_Sub),
                new OleDbParameter("@StopDate",times_End_Sub)
            }; 
            //条码 = @column1 andFormat(测试时间,'yyyy/mm/dd')
           // select cast('2021-06-25 03:02:01' as datetime);



            //var sql = "select * from " + cbx_Model.Text + " where 测试时间  between  #" + times_Start + "#" + " and #" + times_End + "# ";
            var sql = "select * from "+ cbx_Model.Text+ " where CDate(Format(测试时间,'yyyy/MM/dd HH:mm:ss')) between  #" + times_Start_Sub + "#" +" and #"+ Convert.ToDateTime(times_End_Sub) + "# ";
            // 条码,测试人,测试时间,测试结果
            DataTable table = AccessHelper.ExecuteDataTable(conn, sql, pars);
            tablelist();
            dataGridViewDynamic1.SetDataTable(table);

            //foreach (DataRow row in table.Rows)
            //{
            //    foreach (DataColumn column in table.Columns)
            //    {
            //        listBox1.Text+= (row[column] + "\t");
            //    }
            //}
        }
        /// <summary>
        /// 列表
        /// </summary>
        public void tablelist()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            OleDbParameter[] pars1 = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text)
            };
            var sql1 = "select * from " + cbx_Model.Text + " where 条码 = @column1 ";
            DataSet ds = AccessHelper.ExecuteDataSet(conn, sql1, pars1);
            foreach (DataTable tb in ds.Tables)
            {
                foreach (DataColumn col in tb.Columns)
                {
                    dataGridViewDynamic1.AddHeader(col.ColumnName + "\t");
                }
            }


        }

        /// <summary>
        /// 返回datatable
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void datatable()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";

            OleDbParameter[] pars = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text),
                new OleDbParameter("@column2",dateTimePicker1.ToString())
            };
            
            var sql = "select * from HCBI开关组 where 条码 = @column1 and 测试时间 = @column2";
           // 条码,测试人,测试时间,测试结果
            DataTable table = AccessHelper.ExecuteDataTable(conn, sql, pars);
      
                dataGridViewDynamic1.SetDataTable(table);

            //foreach (DataRow row in table.Rows)
            //{
            //    foreach (DataColumn column in table.Columns)
            //    {
            //        listBox1.Text+= (row[column] + "\t");
            //    }
            //}
        }
        /// <summary>
        /// 返回dataSet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void dataset()
        {
            string conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textBoxPath.Text + ";Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";
            OleDbParameter[] pars = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text),
                //new OleDbParameter("@column2",tbx_Vaule.Text)
            };
            var sql = "select 条码,测试人,测试时间 from HCBI开关组 where 条码 = @column1 "; //and 测试节拍 = @column2
            DataTable table = AccessHelper.ExecuteDataTable(conn, sql, pars);
            OleDbParameter[] pars1 = new OleDbParameter[] {
                new OleDbParameter("@column1",textBox_Code.Text)
            };
            var sql1 = "select 条码,测试人,测试时间 from HCBI开关组 where 条码 = @column1 ";
            DataSet ds = AccessHelper.ExecuteDataSet(conn, sql1,pars1);

            foreach (DataTable tb in ds.Tables)
            {
                foreach (DataColumn col in tb.Columns)
                {
                    dataGridViewDynamic1.AddHeader(col.ColumnName + "\t");
                }

                foreach (DataRow row in table.Rows)
                {
                    foreach (DataColumn col in table.Columns)
                    {
                        //(row[col] + "\t");
                    }
                   
                }
            }
        }
        private void buttonSearch_Click(object sender, EventArgs e)
        {
           // DateToAccess();
          dataGridViewDynamic1.ResetGrid(true);
         // datatable();
            chaxun();
           // dataset();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            SaveParameter();
        }
    }
}
