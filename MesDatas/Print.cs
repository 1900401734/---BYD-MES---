using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace MesDatas
{
    public class Print
    {

        private const string appName = "Label Print";
        private const string dataSourced = "Data Sourced";

        private Engine engine = null;
        enum ImageIndex { LoadingFormatImage = 0, FailureToLoadFormatImage = 1 };
        delegate void DelegateShowMessageBox(string message);

        public void Print_Close()
        {
            if (engine != null)
                engine.Stop(SaveOptions.DoNotSaveChanges);
        }
        public List<string> Print_GetPrinter()
        {
            Printers printers = new Printers();
            List<string> printList = new List<string>();
            foreach (Printer printer in printers)
            {
                printList.Add(printer.PrinterName);
            }
            return printList;
        }
        /// <summary>
        /// Bartender打印
        /// </summary>
        /// <param name="printName">打印机名称</param>
        /// <param name="btwPath">BTW文件路径</param>
        /// <param name="key1">模板文件的嵌入名称1</param>
        /// <param name="msg1">传入的变量1</param>
        /// <param name="key2">模板文件的嵌入名称2</param>
        /// <param name="msg2">传入的变量2</param>
        public void Print_P(string printName, string btwPath, string key1, string msg1, string key2, string msg2, out string msg)
        {
            msg = "";
            using (Engine btEngine = new Engine(true))
            {
                if (!File.Exists(btwPath))
                {
                    MessageBox.Show("模板文件不存在");
                    msg = "Fail";
                }
                LabelFormatDocument labelFormat = btEngine.Documents.Open(btwPath);

                try
                {
                    if (msg1.Length != 0)
                        labelFormat.SubStrings.SetSubString(key1, msg1);
                    if (msg2.Length != 0)
                        labelFormat.SubStrings.SetSubString(key2, msg2);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                    msg = "Fail";
                }

                if (printName != "")
                {
                    labelFormat.PrintSetup.PrinterName = printName;
                    labelFormat.Print("BarPrint" + DateTime.Now, 3 * 1000);
                }
                else
                {
                    MessageBox.Show("请先选择打印机", "操作提示");
                }
            }
        }
    }
}
