using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSample
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelHelper.AddColumnNames(new List<string> { "ID", "Name" });
            ExcelHelper.AddRowData(new List<string> { "1", "张三" });
            ExcelHelper.AddRowData(new List<string> { "2", "李四" });
            ExcelHelper.AddRowData(new List<string> { "3", "王武" });
            string saveFileName = @"D:\Temp\text2.xls";
            ExcelHelper.SaveExcel(saveFileName);
        }
    }
}
