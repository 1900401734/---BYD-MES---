using MesDatasCore;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlcCommunication
{
    public partial class FormPlcCommunication : Form
    {
        public FormPlcCommunication()
        {
            InitializeComponent();

        }

        private PlcCommunicationTool cTool = null;
        public void SetTool(cToolBase Tool)
        {
            cTool = (PlcCommunicationTool)Tool;

            InitCombox();
            string spath = Application.StartupPath + "\\ToolParameter\\" + cTool.sName;
            RPLC.SetDataSavePath(spath);
            WPLC.SetDataSavePath(spath);

            RDPLC.SetDataSavePath(spath);
            WDPLC.SetDataSavePath(spath);

            Task.Run(new Action(() =>
            {
                ReadPlcLoop();
            }));
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

            RPLC.LoadCacheDataFromCsv();
            WPLC.LoadCacheDataFromCsv();
            RDPLC.LoadCacheDataFromCsv();
            WDPLC.LoadCacheDataFromCsv();

            textBoxPort.Text = ReadString("textBoxPort");
            textBoxIP.Text = ReadString("textBoxIP");
            comboBoxType.Text = ReadString("ComType");

            //cTool.SetReadPlcActions();
            //cTool.SetWritePlcActions();

        }


        public void SaveParameter()
        {
            // WriteString
            //Invoke(new Action(() =>
            //{
            RPLC.SaveCacheDateToCSV();
            WPLC.SaveCacheDateToCSV();

            RDPLC.SaveCacheDateToCSV();
            WDPLC.SaveCacheDateToCSV();

            WriteString("textBoxPort",  textBoxPort.Text);
            WriteString("textBoxIP", textBoxIP.Text);
            WriteString("ComType", comboBoxType.Text);
            //}));

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("button1_Click");
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            RPLC.Rows.Add();
            MessageBox.Show("button1_Click_1");
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            WPLC.Rows.Add();
        }

        #region "RPLC_ComboBox"

        private ComboBox Score_ComboBox = new ComboBox();
        private void InitCombox()
        {
            comboBoxType.Items.Add("ModBus");
            comboBoxType.Items.Add("FinsTcp");
            comboBoxType.Items.Add("TestPlc");
            //for (int i = 0; i <= 10; i++)
            //{
            //    this.Score_ComboBox.Items.Add(i.ToString());
            //}
            this.Score_ComboBox.Items.Add("Bit");
            this.Score_ComboBox.Items.Add("Byte");
            this.Score_ComboBox.Items.Add("Short");
            this.Score_ComboBox.Text = "0";
            this.Score_ComboBox.Leave += new EventHandler(comboBox50_TextChanged);
            this.Score_ComboBox.SelectedIndexChanged += new EventHandler(comboBox50_TextChanged);
            this.Score_ComboBox.Visible = false;
            this.Score_ComboBox.DropDownStyle = ComboBoxStyle.DropDownList;

            // 将下拉列表框加入到DataGridView控件中
            this.RPLC.Controls.Add(this.Score_ComboBox);
        }

        private void comboBox50_TextChanged(object sender, EventArgs e)
        {
            try
            {
                ((ComboBox)sender).Text = (((ComboBox)sender).Text).ToString();
                int row_index = this.RPLC.CurrentCell.RowIndex;
                this.RPLC.CurrentCell.Value = ((ComboBox)sender).Text;
                this.Score_ComboBox.Visible = false;
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (this.RPLC.CurrentCell != null && this.RPLC.CurrentCell.ColumnIndex == 5/* && this.dataGridView1.CurrentCell.RowIndex < this.dataGridView1.Rows.Count*/)
                {
                    Rectangle rect = RPLC.GetCellDisplayRectangle(RPLC.CurrentCell.ColumnIndex, RPLC.CurrentCell.RowIndex, false);
                    if (RPLC.CurrentCell.Value != null)
                    {
                        string Value = RPLC.CurrentCell.Value.ToString();
                        this.Score_ComboBox.Text = Value;
                        this.Score_ComboBox.Left = rect.Left;
                        this.Score_ComboBox.Top = rect.Top;
                        this.Score_ComboBox.Width = rect.Width;
                        this.Score_ComboBox.Height = rect.Height;
                        this.Score_ComboBox.Visible = true;
                    }

                }
                else
                {
                    this.Score_ComboBox.Visible = false;
                }
            }
            catch (Exception ex)
            {
                throw (ex);
              
            }
        }
        #endregion
        private void ReadPlcLoop()
        {
            while (cTool.IsRunLoop)
            {
                foreach (PlcData cmd in ReadPlcCmd)
                {
                    string plcr = ReadPLC(cmd.addr, cmd.length);
                    if (plcr == cmd.Name)
                    {
                        InteractiveEventArgs eArgs = new InteractiveEventArgs();
                        eArgs.InfoType = InfoType.Command;
                        eArgs.Name = cmd.Name;
                        eArgs.IsKey = false;
                        cTool.SetEnque(eArgs);
                    }
                }
            }

        }
        List<PlcData> ReadPlcCmd = new List<PlcData>();



        private string ReadPLC(string addr, int length)
        {
            return cTool.PlcRead(addr, length);
        }

        private void Contenct_Click(object sender, EventArgs e)
        {
            try
            {
                cTool.PlcConnect(textBoxIP.Text, textBoxPort.Text);
                MessageBox.Show("OK");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void Read_Click(object sender, EventArgs e)
        {
            // cTool.ReadString(ReadPLC)
        
            try
            {
                textBoxReadValue.Text = cTool.PlcRead(textBoxReadAddr.Text, 1);
                MessageBox.Show("OK1");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Write_Click(object sender, EventArgs e)
        {
             cTool.PlcWrite(WriteAddr.Text, textBoxWriteValue.Text);
        }

        private void comboBoxType_SelectedValueChanged(object sender, EventArgs e)
        {
            cTool.SetPlcType(comboBoxType.Text);
        }

        private List<PlcData> GetPlcCmdDatas()
        {
            List<PlcData> ds = new List<PlcData>();
            DataGridView gd = RPLC;
            /******添加行到Datatable控件中*********/
            for (int k = 0; k < gd.Rows.Count; k++)
            {
                if (gd[0, k].Value != null)
                {
                    PlcData data = new PlcData();
                    data.Name = gd[0, k].Value != null ? (string)gd[0, k].Value : "";
                    data.addr = gd[1, k].Value != null ? (string)gd[1, k].Value : "";
                    data.condition = gd[2, k].Value != null ? (string)gd[2, k].Value : "";
                    data.sendCmd = gd[3, k].Value != null ? (string)gd[3, k].Value : "";
                    data.length = gd[4, k].Value != null ? (int)gd[4, k].Value : 0;
                    data.readType = gd[5, k].Value != null ? (ReadType)gd[5, k].Value : 0;
                    data.infoType = InfoType.Command;
                }
            }
            return ds;
        }

        private List<PlcData> GetPlcDatas()
        {
            DataGridView gd = RDPLC;
            List<PlcData> ds = new List<PlcData>();

            /******添加行到Datatable控件中*********/
            for (int k = 0; k < gd.Rows.Count; k++)
            {
                if (gd[0, k].Value != null)
                {
                    PlcData data = new PlcData();
                    data.Name = gd[0, k].Value != null ? (string)gd[0, k].Value : "";
                    data.addr = gd[1, k].Value != null ? (string)gd[1, k].Value : "";
                    data.length = gd[2, k].Value != null ? (int)gd[2, k].Value : 0;
                    data.referenceCmd = gd[3, k].Value != null ? (string)gd[3, k].Value : "";
                    data.sendCmd = gd[4, k].Value != null ? (string)gd[4, k].Value : "";
                    data.readType = gd[5, k].Value != null ? (ReadType)gd[5, k].Value : 0;
                    data.infoType = InfoType.Content;
                }
            }
            return ds;
        }

        private List<PlcData> GetPlcWDatas()
        {
            DataGridView gd = RDPLC;
            List<PlcData> ds = new List<PlcData>();

            /******添加行到Datatable控件中*********/
            for (int k = 0; k < gd.Rows.Count; k++)
            {
                if (gd[0, k].Value != null)
                {
                    PlcData data = new PlcData();
                    data.Name = gd[0, k].Value != null ? (string)gd[0, k].Value : "";
                    data.addr = gd[1, k].Value != null ? (string)gd[1, k].Value : "";
                    data.length = gd[2, k].Value != null ? (int)gd[2, k].Value : 0;
                    data.referenceCmd = gd[3, k].Value != null ? (string)gd[3, k].Value : "";
                    data.sendCmd = gd[4, k].Value != null ? (string)gd[4, k].Value : "";
                    data.readType = gd[5, k].Value != null ? (ReadType)gd[5, k].Value : 0;
                    data.infoType = InfoType.Content;
                }
            }
            return ds;
        }

        private void CmdReadApply_Click_2(object sender, EventArgs e)
        {

            ReadPlcCmd = GetPlcCmdDatas();
        }

        private void DataReadApply_Click_2(object sender, EventArgs e)
        {
            List<PlcData> ds = GetPlcDatas();
            foreach (PlcData p in ds)
            {
                cTool.SetReadPlcActions(p);
            }
        }

        private void CmdWriteApply_Click(object sender, EventArgs e)
        {
            //List<PlcData> ds = GetPlcDatas(WPLC);
            //foreach (PlcData p in ds)
            //{
            //    cTool.SetWritePlcActions(p);
            //}
        }

        private void DataWriteApply_Click(object sender, EventArgs e)
        {
            //List<PlcData> ds = GetPlcDatas(WDPLC);
            //foreach (PlcData p in ds)
            //{
            //    cTool.SetWritePlcActions(p);
            //}
        }
    }

    public class PlcData
    {
        public string addr;
        public int length;
        public ReadType readType;
        public string Name;
        public InfoType infoType;
        public string referenceCmd;
        public string sendCmd;
        public string condition;
    }
    public enum ReadType
    {
        sBit,
        sByte
    }
}
