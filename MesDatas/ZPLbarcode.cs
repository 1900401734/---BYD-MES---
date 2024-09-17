using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
     static class ZPLbarcode
    {
        static string ZPLprinterInit = @"
^XA
~TA000
~JSO
^LT38
^MNW
^MTD
^PON
^PMN
^LH0,0
^JMA
^PR2,2
~SD15
^JUS
^LRN
^CI27
^PA0,1,1,0
^XZ
^XA
^MMT";
        //[Phead]='" + textBox21.Text + "'" +
        //                          ",[Ptime]='" + textBox25.Text + "',[Ptail]='" + textBox26.Text + "'" +
        //                          ",[TFront]='" + textBox20.Text + "',[Tmonarch]='" + textBox27.Text + "'" +
        //                          ",[Tlow]='" + textBox28.Text + "',[Pertow]='" + checkBox2.Checked+ "'" +
        /// <summary>
        /// 单个条码打印机指令
        /// </summary>
        /// <param name="Phead">条码头</param>
        /// <param name="Ptime">条码时间</param>
        /// <param name="Ptail">条码数字</param>
        /// <param name="TFront">文字前端</param>
        /// <param name="Tmonarch">文字后端</param>
        /// <param name="Tlow">文字下端</param>
        /// <returns></returns>
        public static string ZPLsingle(string Phead,string Ptime,string Ptail,string TFront, string Tmonarch,string Tlow)
        {
            string strsing = @"^PW945
^LL236      
^LS0";
            string strTfront = @"^FPH,8^FT47,15^A0R,38,38^FH\^CI28^FD" + TFront + "^FS^CI27";
            string strTmonarch = @"^FPH,3^FT862,118^A0R,40,41^FH\^CI28^FD" + Tmonarch + "^FS^CI27";
            string strTlow = @"^FPH,3^FT306,226^A0N,37,38^FH\^CI28^FD" + Tlow + "^FS^CI27";
            string bacord = @"^BY4,3,105^FT93,144^BCN,,Y,N 
^FH\^FD>;" + Phead + Ptime + Ptail + "^FS";
            string strtals = @"^PQ1,0,1,Y
^ XZ";
            return ZPLprinterInit+"\n"+ strsing + "\n" + strTfront + "\n" + strTmonarch + "\n" + strTlow + "\n" + bacord + "\n" + strtals;
        }
        /// <summary>
        /// 双个条码打印机指令
        /// </summary>
        /// <param name="Phead">条码头</param>
        /// <param name="Ptime">条码时间</param>
        /// <param name="Ptail">条码数字</param>
        /// <param name="TFront">文字前端</param>
        /// <param name="Tmonarch">文字后端</param>
        /// <param name="Tlow">文字下端</param>
        /// <returns></returns>
        public static string ZPLdouble(string Phead, string Ptime, string Ptail, string TFront, string Tmonarch, string Tlow)
        {
            string strsing = @"^PW1913
^LL236
^LS0";
            string strTfront1 = @"^FT1027,70^A0R,38,38^FH\^CI28^FD" + TFront + "^FS^CI27";
            string strTmonarch1 = @"^FT1836,117^A0R,40,41^FH\^CI28^FD" + Tmonarch + "^FS^CI27";
            string strTlow1 = @"^FT1281,224^A0N,37,38^FH\^CI28^FD" + Tlow + "^FS^CI27";
            string bacord1 = @"^BY4,3,105^FT1067,142^BCN,,Y,N 
^FH\^FD>;" + Phead + Ptime + Ptail + "^FS";
            string strTfront2 = @"^FT38,49^A0R,43,43^FH\^CI28^FD" + TFront + "^FS^CI27";
            string strTmonarch2 = @"^FT852,116^A0R,40,41^FH\^CI28^FD" + Tmonarch + "^FS^CI27";
            string strTlow2 = @"^FT297,224^A0N,37,38^FH\^CI28^FD" + Tlow + "^FS^CI27";
            string bacord2 = @"^BY4,3,105^FT84,142^BCN,,Y,N 
^FH\^FD>;" + Phead + Ptime + Ptail + "^FS";
            string pentagram = @"^FO125,145
^GFA,205,688,8,:Z64:eJzVzsENgzAMBdAf5cCt9NZjVugGWaGjdAPYoCtlFEbIkUMU1/5JVUDi3lpITx8bY+Bv6959HlxP8kFXei777OtJrtscHl5wS4hitWCkCd7QCWfaBtP+GNgGB2dbo3Ifx7UmyTTqtmbu/XXX/+SvPNTJi3oZaju/HewL6JAx0QWRJgRznO1R+Qq46CeWr/jJegMIP1t2:7100";
            string strtals = @"^PQ1,0,1,Y
^ XZ";
            return ZPLprinterInit + "\n" + strsing + "\n" + strTfront1 + "\n" + strTmonarch1 + "\n" + strTlow1 + "\n" + bacord1+
                strTfront2 + "\n" + strTmonarch2 + "\n" + strTlow2 + "\n" + bacord2 + "\n" + pentagram + "\n" + strtals;
        }
    }
}
