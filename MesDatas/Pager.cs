using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MesDatas
{
    public class Pager
    {
        public int pageSize { set; get; } = 50;      //每页记录数
        public int recordCount { set; get; } = 0;    //总记录数
        public int pageCount { set; get; } = 0;      //总页数
        public int currentPage { set; get; } = 0;    //当前页
        public DataTable dtSource { set; get; }

        public Pager(DataTable dtSource)
        {
            this.dtSource = dtSource;
        }

        ///LoadPage方法
        /// <summary>
        /// loaddpage方法显示多少页
        /// </summary>
        public DataTable LoadPage()
        {
            if (currentPage < 1) currentPage = 1;
            if (currentPage > pageCount) currentPage = pageCount;

            int beginRecord;
            int endRecord;
            DataTable dtTemp;
            dtTemp = dtSource.Clone();

            beginRecord = pageSize * (currentPage - 1);
            if (currentPage == 1) beginRecord = 0;
            endRecord = pageSize * currentPage;

            if (currentPage == pageCount) endRecord = recordCount;
            if (dtSource.Rows.Count > 0) {
                for (int i = beginRecord; i < endRecord; i++)
                {
                    dtTemp.ImportRow(dtSource.Rows[i]);
                }
            }
            return dtTemp;
        }

        /// <summary>
        /// 分页的方法
        /// </summary>
        /// <param name="str"></param>
        public void fenye()
        {
            recordCount = dtSource.Rows.Count;
            if (pageSize != 0)
            {
                pageCount = (recordCount / pageSize);
                if ((recordCount % pageSize) > 0)
                {
                    pageCount++;
                }
                if (recordCount == 0)
                {
                    pageSize = 0;
                }
            }
            //默认第一页
            currentPage = 1;
        }
    }
}
