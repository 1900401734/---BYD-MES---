using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
namespace MesDatas
{
    public partial class FormCardLogin : Form
    {
        private string m_ConStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Application.StartupPath + @"\用户账号.mdb;Persist Security Info=True;Jet OLEDB:Database Password=byd;User Id=admin";

        private DataSet ds;
        private OleDbDataAdapter dataAdapter;
        public FormCardLogin()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            OleDbConnection sqlCon = new OleDbConnection(m_ConStr);
            string sql = "select * from 用户";
            dataAdapter = new OleDbDataAdapter(sql, sqlCon);
            sqlCon.Close();
            OleDbCommandBuilder cb = new OleDbCommandBuilder(dataAdapter);

            ds = new DataSet();
            dataAdapter.Fill(ds);
            this.dataGridView1.DataSource = ds.Tables[0].DefaultView;

        }
        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            dataAdapter.Update(ds);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            dataAdapter.Update(ds);
        }
    }
}
