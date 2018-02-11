using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MatrixTestResult
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report();

            report.CreateExcelDoc(@"G:\POC\MatrixXl\MatrixTestResult\Report.xlsx");

            Console.WriteLine("Excel file has created!");
        }
    }
}
