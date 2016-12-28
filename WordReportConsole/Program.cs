using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordReportCore;

namespace WordReportConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            ReportFactory factory = new ReportFactory(@"C:\Users\Hyh\Desktop\WordTemplate.docx", @"D:\saveWord.docx");
            factory.BuildWord(new AlarmReport().GerReportView);
        }
    }
}
