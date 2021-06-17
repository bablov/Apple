using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Baebae.Apple.NPOI
{
    public class NPOIHelper
    {
        /// <summary>
        /// DataTable转换EXCEL 文件保存到本地
        /// </summary>
        /// <param name="SourceTable">源数据</param>
        /// <param name="FileName">文件名</param>
        public static void DataTableToExcel(DataTable SourceTable, string FileName)
        {
            using MemoryStream ms = DataTableToExcelStream(SourceTable) as MemoryStream;
            using FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            data = null;
        }
        /// <summary>
        /// DataTable转换EXCEL 文件流
        /// </summary>
        /// <param name="SourceTable">源数据</param>
        /// <param name="DateTimeFormat">时间列格式化</param>
        /// <returns>EXCEL文件流</returns>
        public static Stream DataTableToExcelStream(DataTable SourceTable, string DateTimeFormat = "yyyy-MM-dd HH:mm:ss")
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet();
            IRow headerRow = sheet.CreateRow(0, 20);

            // handling header. 
            //设置顶部大标题样式
            var cellStyleFont = workbook.CreateStyle(
               hAlignment: HorizontalAlignment.Center,
               vAlignment: VerticalAlignment.Center,
               fontHeightInPoints: 12,
               isBold: true,
               isAddBorder: true,
               isAddCellBackground: true,
               cellBackgroundColor: HSSFColor.Gold.Index,
               fillPattern: FillPattern.SolidForeground,
               fontColor: HSSFColor.Black.Index);

            foreach (DataColumn column in SourceTable.Columns)
            {
                headerRow.CreateCells(cellStyleFont, column.Ordinal, column.ColumnName);
            }
            // handling value. 
            int rowIndex = 1;
            //单元格边框样式
            var cellStyle = workbook.CreateStyle(
                hAlignment: HorizontalAlignment.Center,
                vAlignment: VerticalAlignment.Center,
                fontHeightInPoints: 10,
                isAddBorder: true);

            foreach (DataRow row in SourceTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                IDataFormat dataformat = workbook.CreateDataFormat();

                foreach (DataColumn column in SourceTable.Columns)
                {
                    if (row[column] is DBNull)
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(string.Empty);
                        continue;
                    }

                    if (column.DataType == typeof(int))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((int)row[column]);
                    }
                    else if (column.DataType == typeof(float))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((float)row[column]);
                    }
                    else if (column.DataType == typeof(double))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((double)row[column]);
                    }
                    else if (column.DataType == typeof(Byte))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((byte)row[column]);
                    }
                    else if (column.DataType == typeof(UInt16))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((UInt16)row[column]);
                    }
                    else if (column.DataType == typeof(UInt32))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((UInt32)row[column]);
                    }
                    else if (column.DataType == typeof(UInt64))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((UInt64)row[column]);
                    }
                    else if (column.DataType == typeof(DateTime))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((DateTime)row[column]);
                        cellStyle.DataFormat = dataformat.GetFormat(DateTimeFormat);
                        dataRow.GetCell(column.Ordinal).CellStyle = cellStyle;
                    }
                    else
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(Convert.ToString(row[column]));
                        dataRow.GetCell(column.Ordinal).CellStyle = cellStyle;
                    }
                }
                rowIndex++;
            }

            workbook.Write(ms);
            ms.Flush();
            //ms.Position = 0;

            //sheet = null;
            //headerRow = null;
            //workbook = null;

            return ms;
        }

        /// <summary>
        /// DataTable转换EXCEL 文件流 (行数超过65535，sheet分页)
        /// </summary>
        /// <param name="SourceTable">源数据</param>
        /// <param name="sheetSize">sheet最大行数，不大于65535</param>
        /// <param name="DateTimeFormat">时间列格式化</param>
        /// <returns>EXCEL文件流</returns>
        public static Stream DataTableToExcelStreamPage(DataTable SourceTable, int sheetSize = 65535, string DateTimeFormat = "yyyy-MM-dd HH:mm:ss")
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();

            IDataFormat dataformat = workbook.CreateDataFormat();

            int count = SourceTable.Rows.Count;
            int total = count / sheetSize + (count % sheetSize > 0 ? 1 : 0);

            for (int sheetIndex = 0; sheetIndex < total; sheetIndex++)
            {
                ISheet sheet = workbook.CreateSheet();
                IRow headerRow = sheet.CreateRow(0, 20);
                // handling header. 
                //设置顶部大标题样式
                var cellStyleFont = workbook.CreateStyle(
                   hAlignment: HorizontalAlignment.Center,
                   vAlignment: VerticalAlignment.Center,
                   fontHeightInPoints: 12,
                   isBold: true,
                   isAddBorder: true,
                   isAddCellBackground: true,
                   cellBackgroundColor: HSSFColor.Gold.Index,
                   fillPattern: FillPattern.SolidForeground,
                   fontColor: HSSFColor.Black.Index);

                foreach (DataColumn column in SourceTable.Columns)
                    headerRow.CreateCells(cellStyleFont, column.Ordinal, column.ColumnName);


                // handling value. 
                int rowIndex = 1;
                //单元格边框样式
                var cellStyle = workbook.CreateStyle(
                    hAlignment: HorizontalAlignment.Center,
                    vAlignment: VerticalAlignment.Center,
                    fontHeightInPoints: 10,
                    isAddBorder: true);
                for (int i = sheetIndex * sheetSize; i < (total.Equals(sheetIndex + 1) ? count : (sheetIndex + 1) * sheetSize); i++)
                {
                    DataRow row = SourceTable.Rows[i];

                    IRow dataRow = sheet.CreateRow(rowIndex);

                    foreach (DataColumn column in SourceTable.Columns)
                    {
                        if (row[column] is DBNull)
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(string.Empty);
                            continue;
                        }
                        if (column.DataType == typeof(int))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((int)row[column]);
                        }
                        else if (column.DataType == typeof(float))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((float)row[column]);
                        }
                        else if (column.DataType == typeof(double))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((double)row[column]);
                        }
                        else if (column.DataType == typeof(Byte))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((byte)row[column]);
                        }
                        else if (column.DataType == typeof(UInt16))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((UInt16)row[column]);
                        }
                        else if (column.DataType == typeof(UInt32))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((UInt32)row[column]);
                        }
                        else if (column.DataType == typeof(UInt64))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((UInt64)row[column]);
                        }
                        else if (column.DataType == typeof(DateTime))
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue((DateTime)row[column]);
                            cellStyle.DataFormat = dataformat.GetFormat(DateTimeFormat);
                            dataRow.GetCell(column.Ordinal).CellStyle = cellStyle;
                        }
                        else
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(Convert.ToString(row[column]));
                            dataRow.GetCell(column.Ordinal).CellStyle = cellStyle;
                        }
                    }
                    rowIndex++;
                }

                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                sheet = null;
                headerRow = null;
            }
            workbook = null;

            return ms;
        }
        /// <summary>
        /// EXCEL文件 转换成 DataTable
        /// </summary>
        /// <param name="ExcelFileStream">EXCEL文件流</param>
        /// <returns></returns>
        public static DataTable ExcelFileToDataTable(string file,string sheetname)
        {
            string fileExt = Path.GetExtension(file);
            IWorkbook workbook = null;
            if (fileExt == ".xls")
                workbook = new HSSFWorkbook(FileToStream(file));
            else if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook(new FileInfo(file));


            ISheet sheet = workbook.GetSheet(sheetname);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }
            for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    dataRow[j] = row.GetCell(j).ToString();
                    if(string.IsNullOrEmpty(dataRow[j].ToString()) && j > 0)
                    {
                        dataRow[j] = dataRow[j-1].ToString();
                    }
                }
                table.Rows.Add(dataRow);
            }


            workbook = null;
            sheet = null;
            return table;
        }
        public static DataSet ExcelFileToDataSet(string file)
        {
            DataSet dataSet = new DataSet();
            string fileExt = Path.GetExtension(file);
            IWorkbook workbook = null;
            if (fileExt == ".xls")
                workbook = new HSSFWorkbook(FileToStream(file));
            else if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook(new FileInfo(file));

            var sheets = workbook.NumberOfSheets;
            for (int s = 0; s < sheets; s++)
            {
                ISheet sheet = workbook.GetSheetAt(s);
                DataTable table = new DataTable(sheet.SheetName.Trim());
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue.Trim());
                    table.Columns.Add(column);
                }
                for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dataRow = table.NewRow();

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        dataRow[j] = row.GetCell(j).ToString();
                    table.Rows.Add(dataRow);
                }
                sheet = null;
                dataSet.Tables.Add(table);
            }
            workbook = null;
            return dataSet;
        }

        /// <summary>
        /// EXCEL文件流 转换成 DataTable
        /// </summary>
        /// <param name="ExcelFileStream">EXCEL文件流</param>
        /// <returns></returns>
        public static DataTable ExcelStreamToDataTable(Stream ExcelFileStream, string fileExt)
        {
            IWorkbook workbook = null;
            if (fileExt == ".xls")
                workbook = new HSSFWorkbook(ExcelFileStream);
            else if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook(ExcelFileStream);

            ISheet sheet = workbook.GetSheetAt(0);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue.Trim());
                table.Columns.Add(column);
            }

            for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                    dataRow[j] = row.GetCell(j).ToString();
                table.Rows.Add(dataRow);
            }

            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return table;
        }
        public static DataSet ExcelStreamToDataSet(Stream ExcelFileStream, string fileExt)
        {
            DataSet dataSet = new DataSet();
            IWorkbook workbook = null;
            if (fileExt == ".xls")
                workbook = new HSSFWorkbook(ExcelFileStream);
            else if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook(ExcelFileStream);

            var sheets = workbook.NumberOfSheets;
            for (int s = 0; s < sheets; s++)
            {
                ISheet sheet = workbook.GetSheetAt(s);
                DataTable table = new DataTable(sheet.SheetName.Trim());
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue.Trim());
                    table.Columns.Add(column);
                }
                for (int i = (sheet.FirstRowNum + 1); i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dataRow = table.NewRow();

                    for (int j = row.FirstCellNum; j < cellCount; j++)
                        dataRow[j] = row.GetCell(j).ToString();
                    table.Rows.Add(dataRow);
                }

                ExcelFileStream.Close();
                sheet = null;
                dataSet.Tables.Add(table);
            }
            workbook = null;
            return dataSet;
        }
        /// <summary>
        /// EXCEL文件流 转换成 DataTable
        /// </summary>
        /// <param name="ExcelFileStream"></param>
        /// <param name="file"></param>
        /// <param name="SheetIndex"></param>
        /// <param name="HeaderRowIndex"></param>
        /// <returns></returns>
        public static DataTable ExcelStreamToDataTable(Stream ExcelFileStream, string file, int SheetIndex, int HeaderRowIndex)
        {
            string fileExt = Path.GetExtension(file);
            IWorkbook workbook = null;
            if (fileExt == ".xls")
                workbook = new HSSFWorkbook(ExcelFileStream);
            else if (fileExt == ".xlsx")
                workbook = new XSSFWorkbook(ExcelFileStream);

            ISheet sheet = workbook.GetSheetAt(SheetIndex);
            DataTable table = new DataTable();
            IRow headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;

            for (int i = 0; i < rowCount + 1; i++)
            {

                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();
                for (int j = row.FirstCellNum; j <= cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();
                }
                table.Rows.Add(dataRow);
            }
            table.Rows.RemoveAt(0);
            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return table;
        }

        public static void GenericToExcel(Dictionary<string, List<string>> Source, string sheetname, string FileName)
        {
            using MemoryStream ms = GenericToExcelStream(Source, sheetname) as MemoryStream;
            using FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();
            fs.Write(data, 0, data.Length);
            fs.Flush();
            data = null;
        }

        public static Stream GenericToExcelStream(Dictionary<string, List<string>> Source, string sheetname)
        {
            IWorkbook workbook = new XSSFWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet(sheetname);
            IRow headerRow = sheet.CreateRow(0, 20);


            var headers = Source.Select(o => o.Key).ToList();
            var values = Source.Select(o => o.Value).ToList();
            var valueCount = values.Max(o => o.Count);
            // handling header. 
            //设置顶部大标题样式
            var cellStyleFont = workbook.CreateStyle(
               hAlignment: HorizontalAlignment.Center,
               vAlignment: VerticalAlignment.Center,
               fontHeightInPoints: 12,
               isBold: true,
               isAddBorder: true,
               isAddCellBackground: true,
               cellBackgroundColor: HSSFColor.Gold.Index,
               fillPattern: FillPattern.SolidForeground,
               fontColor: HSSFColor.Black.Index);

            for (int i = 0; i < headers.Count; i++)
            {
                headerRow.CreateCells(cellStyleFont, i, headers[i]);
            }
            // handling value. 

            //单元格边框样式
            var cellStyle = workbook.CreateStyle(
                hAlignment: HorizontalAlignment.Justify,
                vAlignment: VerticalAlignment.Center,
                fontHeightInPoints: 10,
                isAddBorder: true,
                isLineFeed: false);

            for (int i = 0; i < valueCount; i++)
            {
                IRow row2 = sheet.CreateRow(i + 1);
                for (int j = 0; j < values.Count; j++)
                {
                    if (values[j].Count > i)
                    {
                        row2.CreateCell(j).SetCellValue(values[j][i]);
                        row2.GetCell(j).CellStyle = cellStyle;
                    }
                    else
                        row2.CreateCell(j).SetCellValue("");
                }
            }


            workbook.Write(ms);
            ms.Flush();
            //ms.Position = 0;

            //sheet = null;
            //headerRow = null;
            //workbook = null;

            return ms;
        }
        static Stream FileToStream(string fileName)
        {
            // 打开文件
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
            // 读取文件的 byte[]
            byte[] bytes = new byte[fileStream.Length];
            fileStream.Read(bytes, 0, bytes.Length);
            fileStream.Close();
            // 把 byte[] 转换成 Stream
            Stream stream = new MemoryStream(bytes);
            return stream;
        }
    }
}
