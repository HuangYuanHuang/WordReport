using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportCore
{
    public class ReportView
    {
        /// <summary>
        /// Word 文本域 MergeField 字段字典
        /// </summary>
        public Dictionary<string, object> MergeFields { get; set; } = new Dictionary<string, object>();

        public List<TableModel> Tables { get; set; } = new List<TableModel>();

        public List<ChartModel> Charts { get; set; } = new List<ChartModel>();
    

    }


}
