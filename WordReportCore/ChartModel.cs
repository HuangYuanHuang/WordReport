using Aspose.Words.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportCore
{
    public class ChartModel
    {
        /// <summary>
        /// WORD 书签
        /// </summary>
        public string BookMark { get; set; }

        public string ChartTitle { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }
        public ChartType Type { get; set; }
        public List<ChartNodeModel> Nodes { get; set; }

        public ChartModel()
        {
            Nodes = new List<ChartNodeModel>();
            Width = 432;
            Height = 260;
            Type = ChartType.Column;
        }


    }
    public class ChartNodeModel
    {
        public string SeriesName { get; set; }

        public DateTime[] XAxisDate { get; set; }

        public string[] XAxisStrings { get; set; }

        public double[] YaxisValues { get; set; }
    }
}
