using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatasCore
{

    /// <summary>
    /// 将操作封装为对象
    /// </summary>
    public class ActionBase
    {
        public cToolBase cTool;
        public virtual void Run()
        {
            HandlerActionWithArgus?.Invoke(eArgus);
        }

        public event ActionHandlerWithArgus HandlerActionWithArgus;

        private BaseActionArgus eArgus;

        public void SetArgus(BaseActionArgus e)
        {
            eArgus = e;
        }
       
    }

    public class BaseActionArgus : EventArgs
    {
        public int ID { get; set; }
        public string ParameterName { get; set; }
        public BaseActionArgus()
        {

        }
    }

    public delegate int ActionHandlerWithArgus(BaseActionArgus argus);
}
