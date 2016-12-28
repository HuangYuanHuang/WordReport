using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using System.Data;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing;

namespace WordReportCore
{
    public class ReportFactory
    {
        string savePath;
        Document docReport;
        DocumentBuilder builder;
        ReportView reportView;
        public ReportFactory(string templatePath, string savePath)
        {
            this.savePath = savePath;
            docReport = new Document(templatePath);
            builder = new DocumentBuilder(docReport);
        }
        public void BuildWord(ReportView report)
        {
            reportView = report;
            BuildFiled();
            BuildForEachTable();
            BuildChart();
            docReport.Save(savePath);

        }

        private void BuildFiled()
        {
            var fieldNames = reportView.MergeFields.Keys.ToArray();
            var fieldValues = reportView.MergeFields.Values.ToArray();
            docReport.MailMerge.Execute(fieldNames, fieldValues);
        }


        private void BuildChart()
        {



            foreach (var item in reportView.Charts)
            {
                builder.MoveToBookmark(item.BookMark);
                Shape shape = builder.InsertChart(item.Type, item.Width, item.Height);
                Chart chart = shape.Chart;
                chart.Title.Text = item.ChartTitle;
                ChartSeriesCollection seriesColl = chart.Series;
                seriesColl.Clear();
                item.Nodes.ForEach(d =>
                {
                    if (d.XAxisDate?.Length > 0)
                    {
                        seriesColl.Add(d.SeriesName, d.XAxisDate, d.YaxisValues);
                    }
                    else if (d.XAxisStrings?.Length > 0)
                    {
                        seriesColl.Add(d.SeriesName, d.XAxisStrings, d.YaxisValues);
                    }

                });
            }


        }
        private void BuildForEachTable()
        {

            foreach (var item in reportView.Tables)
            {
                DataTable datatable = new DataTable(item.ListName);
                item.ColumsName.ToList().ForEach(d => datatable.Columns.Add(d));
                item.Rows.ForEach(d =>
                {
                    var row = datatable.NewRow();
                    d.ForEach(column =>
                    {
                        row[column.ColumnName] = column.Value;
                    });
                    datatable.Rows.Add(row);

                });
                docReport.MailMerge.ExecuteWithRegions(datatable);
            }


        }
    }
}
