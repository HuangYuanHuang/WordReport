using Aspose.Words.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordReportCore
{
    public class AlarmReport
    {
        ReportView mainReport = new ReportView();

        public ReportView GerReportView { get { return mainReport; } }
        public AlarmReport()
        {
            InitFiled();
            InitTable();
            InitChart();
        }
        public void InitFiled()
        {
            DateTime time = DateTime.Now.AddMonths(-1);
            string timeSpan = $"{time.Year}年{time.Month}月1日至{time.Year}年{time.Month}月{DateTime.DaysInMonth(time.Year, time.Month)}日";

            mainReport.MergeFields.Add("ReportTitle", "5#催化裂化装置评估报告");
            mainReport.MergeFields.Add("TimeSpan", timeSpan);


        }

        public void InitTable()
        {
            //震颤报警
            TableModel tableModel = new TableModel()
            {
                ListName = "ChatterList",
                ColumsName = new string[] { "Tag", "AlarmID", "Priority", "ClusterMember", "NumberOfClusters" }
            };
            mainReport.Tables.Add(tableModel);
            for (int i = 0; i < 20; i++)
            {
                List<TableColumnNode> colums = new List<TableColumnNode>()
                {
                    new TableColumnNode() { ColumnName="Tag",Value="T123pc"},
                    new TableColumnNode() { ColumnName="AlarmID",Value="BADPV"},
                    new TableColumnNode() { ColumnName="Priority",Value="LOW"},
                    new TableColumnNode() { ColumnName="ClusterMember",Value="95.74%"},
                    new TableColumnNode() { ColumnName="NumberOfClusters",Value="90"},

                };
                tableModel.Rows.Add(colums);

            }

            //主要KPI 
            tableModel = new TableModel()
            {
                ListName = "MainKpi",
                ColumsName = new string[] { "Title", "PreValue", "CurrentValue", "Perent", "Mark" }
            };
            mainReport.Tables.Add(tableModel);
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="高峰报警数(每十分钟×每个内操)"},
                new TableColumnNode() { ColumnName="PreValue",Value=16.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=13.31},
                new TableColumnNode() { ColumnName="Perent",Value=-18.21},
                new TableColumnNode() { ColumnName="Mark",Value="<15"}
            });
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="平均报警数(每十分钟×每个内操)"},
                new TableColumnNode() { ColumnName="PreValue",Value=16.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=13.31},
                new TableColumnNode() { ColumnName="Perent",Value=-18.21},
                new TableColumnNode() { ColumnName="Mark",Value="<2"}
            });
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="扰动比例%"},
                new TableColumnNode() { ColumnName="PreValue",Value=16.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=13.31},
                new TableColumnNode() { ColumnName="Perent",Value=-18.21},
                new TableColumnNode() { ColumnName="Mark",Value="<2"}
            });

            ////震颤报警
            tableModel = new TableModel()
            {
                ListName = "ChattList",
                ColumsName = new string[] { "Title", "PreValue", "CurrentValue", "Perent" }
            };
            mainReport.Tables.Add(tableModel);
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="震颤报警总数(前20位)个"},
                new TableColumnNode() { ColumnName="PreValue",Value=16.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=13.31},
                new TableColumnNode() { ColumnName="Perent",Value=-18.21},

            });
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="震颤报警总数(前20位)占总报警数百分比"},
                new TableColumnNode() { ColumnName="PreValue",Value=46.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=23.31},
                new TableColumnNode() { ColumnName="Perent",Value=+18.21},

            });

            //最频繁报警
            tableModel = new TableModel()
            {
                ListName = "MostAlarmList",
                ColumsName = new string[] { "Title", "PreValue", "CurrentValue", "Perent" }
            };
            mainReport.Tables.Add(tableModel);
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="最频繁报警总数(前20位)个"},
                new TableColumnNode() { ColumnName="PreValue",Value=16.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=13.31},
                new TableColumnNode() { ColumnName="Perent",Value=-18.21},

            });
            tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="Title",Value="最频繁报警总数(前20位)占总报警数百分比"},
                new TableColumnNode() { ColumnName="PreValue",Value=46.26},
                new TableColumnNode() { ColumnName="CurrentValue",Value=23.31},
                new TableColumnNode() { ColumnName="Perent",Value=+18.21},

            });

            //因果报警
            tableModel = new TableModel()
            {
                ListName = "SymptomaticList",
                ColumsName = new string[] { "ParentTag", "ParentAlarmIdentifier", "ChildTag", "ChildAlarmIdentifier" , "OccurrenceCount", "Predictability" , "Significance" }
            };
            mainReport.Tables.Add(tableModel);

            for (int i = 0; i < 20; i++)
            {
                tableModel.Rows.Add(new List<TableColumnNode>()
            {
                new TableColumnNode() { ColumnName="ParentTag",Value="LI93103"},
                new TableColumnNode() { ColumnName="ParentAlarmIdentifier",Value="PVHI"},
                new TableColumnNode() { ColumnName="ChildTag",Value="ZSO93103"},
                new TableColumnNode() { ColumnName="ChildAlarmIdentifier",Value="PVLO"},
                new TableColumnNode() { ColumnName="OccurrenceCount",Value=167},
                new TableColumnNode() { ColumnName="Predictability",Value="75.12%"},
                new TableColumnNode() { ColumnName="Significance",Value="65.43%"},

            });
            }
        }

        public void InitChart()
        {
            ChartModel chartModel = new ChartModel() { BookMark = "Performance", ChartTitle = "Alarm Performance Over Time", Type = ChartType.Column };
            mainReport.Charts.Add(chartModel);
            List<DateTime> list = new List<DateTime>();
            List<double> listValue = new List<double>();
            for (int i = 0; i < 31; i++)
            {
                list.Add(DateTime.Now.Date.AddDays(i));

            }
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "XTB2#FT", XAxisDate = list.ToArray(), YaxisValues = (Enumerable.Range(1, 31).OrderBy(d => Guid.NewGuid()).Select(d => Convert.ToDouble(d)).ToArray()) });
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "FTB2#FT", XAxisDate = list.ToArray(), YaxisValues = (Enumerable.Range(1, 31).OrderBy(d => Guid.NewGuid()).Select(d => Convert.ToDouble(d)).ToArray()) });


            Action<string, string, ChartType, string> actionChart = (mark, title, type, name) =>
              {
                  chartModel = new ChartModel() { BookMark = mark, ChartTitle = title, Type = type };
                  mainReport.Charts.Add(chartModel);
                  chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = name, XAxisDate = list.ToArray(), YaxisValues = (Enumerable.Range(1, 31).OrderBy(d => Guid.NewGuid()).Select(d => Convert.ToDouble(d)).ToArray()) });

              };
            actionChart("AverageAlarmRate", "Average Alarm Rate Over Time", ChartType.LineStacked, "XTB2#FT");
            actionChart("AverageAlarmHis", "平均报警数历史趋势", ChartType.LineStacked, "XTB2#FT");

            actionChart("MaxAlarmRate", "Maximum Alarm Rate Over Time", ChartType.LineStacked, "XTB2#FT");
            actionChart("MaxAlarmRateHis", "最大报警数历史趋势", ChartType.LineStacked, "XTB2#FT");

            actionChart("PercentUpset", "Percentage of 10 Minutes in Burst Over Time", ChartType.LineStacked, "XTB2#FT");
            actionChart("PercentUpsetHis", "扰动比例历史趋势", ChartType.LineStacked, "XTB2#FT");

            chartModel = new ChartModel() { BookMark = "AlarmCount", ChartTitle = "报警计数统计", Type = ChartType.Line };
            mainReport.Charts.Add(chartModel);
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "报警数", XAxisDate = list.ToArray(), YaxisValues = (Enumerable.Range(1, 31).OrderBy(d => Guid.NewGuid()).Select(d => Convert.ToDouble(d)).ToArray()) });
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "目标值", XAxisDate = list.ToArray(), YaxisValues = Enumerable.Range(1, 31).Select(d => Convert.ToDouble(14)).ToArray() });
            actionChart("AlarmCountHis", "报警数历史趋势", ChartType.LineStacked, "报警数");

            chartModel = new ChartModel() { BookMark = "AlarmArea", ChartTitle = "报警区域分布统计", Type = ChartType.Pie };
            mainReport.Charts.Add(chartModel);
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "区域百分比", XAxisStrings = new string[] { "023", "024", "025", "GDS" }, YaxisValues = new double[] { 0.2, 0.2, 0.5, 0.1 } });

            actionChart("MostAlarm", "最频繁报警分布（前20位）",ChartType.Column,"Identifier");

            chartModel = new ChartModel() { BookMark = "AlarmPrority", ChartTitle = "报警优先级分布统计", Type = ChartType.Pie };
            mainReport.Charts.Add(chartModel);
            chartModel.Nodes.Add(new ChartNodeModel() { SeriesName = "优先级百分比", XAxisStrings = new string[] { "LOW", "HIGH", "URGENT" }, YaxisValues = new double[] { 0.2, 0.3, 0.5 } });
        }
    }
}
