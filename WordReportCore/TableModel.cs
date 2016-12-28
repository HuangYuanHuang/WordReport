using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportCore
{
    public class TableModel
    {
        /// <summary>
        /// word模板 tablestart 名称
        /// </summary>
        public string ListName { get; set; }

        /// <summary>
        /// 列集合
        /// </summary>
        public string[] ColumsName { get; set; }

        public List<List<TableColumnNode>> Rows { get; set; }

        public TableModel()
        {
            Rows = new List<List<TableColumnNode>>();
        }
    }

    public class TableColumnNode
    {
        public string ColumnName { get; set; }

        public object Value { get; set; }
    }
}
