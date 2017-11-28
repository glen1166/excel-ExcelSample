using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSample
{
    public class ExcelHelper
    {
        private static DataTable _data = new DataTable();

        public static void AddColumnNames(List<string> columnNames)
        {
            _data.Clear();
            foreach (var item in columnNames)
            {
                _data.Columns.Add(item);
            }
        }

        public static void AddRowData(List<string> data)
        {
            DataRow row = _data.NewRow();
            for (int i = 0; i < _data.Columns.Count; i++)
            {
                row[_data.Columns[i]] = data[i];
            }

            _data.Rows.Add(row);
        }

        public static void ClearData()
        {
            _data.Clear();
        }

        public static Stream RenderDataTableToExcel()
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet();
            IRow headerRow = sheet.CreateRow(0);

            // handling header. 
            foreach (DataColumn column in _data.Columns)
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            // handling value. 
            int rowIndex = 1;

            foreach (DataRow row in _data.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in _data.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }

                rowIndex++;
            }

            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            sheet = null;
            headerRow = null;
            workbook = null;

            return ms;
        }

        public static void CopyToMemory(Stream input, FileStream fileStream)
        {
            byte[] buffer = new byte[8192];
            int bytesRead;
            while ((bytesRead = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                fileStream.Write(buffer, 0, bytesRead);
            }
            // Rewind ready for reading (typical scenario)
            fileStream.Position = 0;
        }

        public static void SaveExcel(string filePath)
        {
            try
            {
                var ms = RenderDataTableToExcel();
                using (FileStream fs = File.Create(filePath))
                {
                    CopyToMemory(ms, fs);
                }

                ms.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("导出文件时出错,文件可能正被打开！\n" + ex.Message);
            }
        }
    }
}
