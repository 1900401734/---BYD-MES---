using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas.Utiey
{
    class TypeRead
    {
        public static string Typerdess(string redss,string typelen)
        {
            string bb = "";
            if (typelen.Equals("0"))
            {
                bb = CodeNum.PNumCode(redss);
            }
            else if(typelen.Equals("1")){
                bb = CodeNum.PNtimeCode(redss);
                
            }else if (typelen.Equals("2"))
            {
                bb = CodeNum.PdounInCode(redss);
            }else if (typelen.Equals("3"))
            {
                bb = CodeNum.PNumOKAG(redss);
            }
            else if(typelen.Equals("4"))
            {
                bb = redss;
            }
            else if (typelen.Equals("5"))
            {
                bb = CodeNum.PThousdCode(redss);
            }
            else if (typelen.Equals("6"))
            {
                bb = CodeNum.PThouInCode(redss);
            }
            return bb;
        }
    }
}
