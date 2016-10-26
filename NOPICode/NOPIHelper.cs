using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using NPOI.HSSF.EventUserModel;

namespace NOPICode
{
    public class NOPIHelper
    {
        /// <summary>
        /// 导出出指定数据类型的excel
        /// </summary>
        /// <typeparam name="T">数据源类型</typeparam>
        /// <param name="headers">标题栏 可支持多行以及单元格合并</param>
        /// <param name="datas">数据源</param>
        /// <param name="filterData">过滤数据</param>
        /// <param name="sheetName">工作簿名字</param>
        /// <returns></returns>
        public static byte[] Export<T>(string[,] headers, IEnumerable<T> datas, Func<T, object> filterData, string sheetName = "sheet1")
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            ISheet sheet = workbook.CreateSheet(sheetName);
            int i = 0;
            int starRowMerge = -1;
            int starCellMerge = -1;
            int endCellMerge = -1;
            for (i = 0; i < headers.GetLength(0); i++)
            {
                if (starRowMerge <= 0)
                {
                    starRowMerge = i;//开始合并行
                }
                var row = sheet.CreateRow(i);
                for (int j = 0; j < headers.GetLength(1); j++)
                {
                    var cell = row.CreateCell(j);
                    var value = headers[i, j];
                    if (string.IsNullOrEmpty(value))
                    {
                        if (j == headers.GetLength(1) - 1)
                        {
                           var cellIndex=  sheet.AddMergedRegion(new CellRangeAddress(starRowMerge, starRowMerge, starCellMerge, j));
                           var newCell = row.GetCell(starCellMerge);
                           SetCellCenter(newCell,workbook.CreateCellStyle());
                            endCellMerge = -1;
                        }
                        else
                        {
                            endCellMerge = j;
                        }
                    }
                    else
                    {
                        if (endCellMerge >= 0)
                        {
                            //合并列
                          var cellIndex=sheet.AddMergedRegion(new CellRangeAddress(starRowMerge, starRowMerge, starCellMerge, endCellMerge));
                            endCellMerge = -1;
                            starRowMerge = 0;
                            var newCell = row.GetCell(starCellMerge);
                            SetCellCenter(newCell,workbook.CreateCellStyle());
                        }
                        starCellMerge = j;
                        cell.SetCellValue(value);
                    }
                }
            }
            object exportData = null;
            foreach (var data in datas)
            {
                var row = sheet.CreateRow(i++);
                exportData = data;
                if (filterData != null)
                {
                    exportData = filterData(data);
                }
                row.CreateCells(exportData);
            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Seek(0, SeekOrigin.Begin);
                byte[] bytedatas = new byte[ms.Length];
                ms.Read(bytedatas, 0, bytedatas.Length);
                return bytedatas;
            }
        }


        /// <summary>
        /// 通过模板导出数据
        /// </summary>
        /// <typeparam name="T">数据源类型</typeparam>
        /// <param name="path">需要绝对路径</param>
        /// <param name="dataTitleRowIndex">模板数据标题列所在行()</param>
        /// <param name="replaceDic">替换模板的标示符</param>
        /// <param name="datas">数据源</param>
        /// <param name="count">数据源总条数</param>
        /// <param name="filterData">过滤数据</param>
        /// <param name="sheetName">工作簿名字</param>
        /// <returns></returns>
        public static byte[] Export<T>(string path,int dataTitleRowIndex, Dictionary<string, string> replaceDic, IEnumerable<T> datas,int count, Func<T, object> filterData,
            string sheetName = "sheet1")
        {
            string extension= Path.GetExtension(path);
            if (!extension.Equals(".xls") && !extension.Equals(".xlsx"))
            {
                throw new FileLoadException("文件格式错误");
            }
            if (!File.Exists(path))
            {
                throw new FileLoadException("文件不存在");
            }
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook book = new HSSFWorkbook(stream);
                ISheet sheet=book.GetSheetAt(0);
                //先执行替换操作
                var rows = sheet.GetEnumerator();
                while (rows.MoveNext())
                {
                    var cuurentRow = (HSSFRow)rows.Current;
                    ReplaceCell(cuurentRow, replaceDic);
                }
                //如果底部有替换行 将底部的行 移动到插入数据之后
                var rowCount = sheet.LastRowNum;
                if (sheet.LastRowNum > dataTitleRowIndex)
                {
                    for (int i = dataTitleRowIndex + 1; i <=rowCount; i++)
                    {

                        sheet.MoveCell(i, i + count);

                    }
                }
                int j = dataTitleRowIndex + 1;
                object exportData = null;
                foreach (var data in datas)
                {
                   
                    var row = sheet.GetRow(j);
                    if (row == null)
                    {
                        row = sheet.CreateRow(j);
                    }
                    j++;
                    exportData = data;
                    if (filterData != null)
                    {
                        exportData = filterData(data);
                    }
                    row.CreateCells(exportData);
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    book.Write(ms);
                    ms.Seek(0, SeekOrigin.Begin);
                    byte[] bytedatas = new byte[ms.Length];
                    ms.Read(bytedatas, 0, bytedatas.Length);
                    return bytedatas;
                }
            }
        }

        /// <summary>
        /// 替换表格中的占位符
        /// </summary>
        /// <param name="row"></param>
        /// <param name="replaceDic"></param>
        private static void ReplaceCell(IRow row, Dictionary<string, string> replaceDic)
        {
           var cellsLength=row.LastCellNum;
            for (int i = 0; i < cellsLength; i++)
            {
                var cell = row.GetCell(i);
                var value= cell.StringCellValue;
                if (replaceDic.ContainsKey(value))
                {
                    cell.SetCellValue(replaceDic[value]);
                }
            }
        }

        //设置单元格水平垂直居中
        private static void SetCellCenter(ICell cell,ICellStyle cellStyle)
        {
            cellStyle.VerticalAlignment = VerticalAlignment.JUSTIFY;//垂直对齐(默认应该为center，如果center无效则用justify)
            cellStyle.Alignment = HorizontalAlignment.CENTER;//水平对齐
            cell.CellStyle = cellStyle;
        }


        /// <summary>
        /// 将excel转换为datable
        /// </summary>
        /// <param name="fileStream">excel文件流</param>
        /// <param name="skip">忽略处理列数</param>
        /// <returns></returns>
        public static DataTable GetDataTable(Stream fileStream, int skip=0)
        {
            DataTable dt = new DataTable();
            try
            {
                if (fileStream == null) return dt;
                IWorkbook workbook = new HSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0);
                IRow firstRow = sheet.GetRow(0);
                if (firstRow == null) return dt; //空Excel文件
                int coluNum = firstRow.LastCellNum;//列数
                //根据excel的列数定义列
                for (int colInx = 0; colInx < coluNum; colInx++)
                  dt.Columns.Add(Convert.ToChar(((int)'A') + colInx).ToString());
                int startRowInx=sheet.FirstRowNum+skip;
                for (int rowInx = startRowInx; rowInx <= sheet.LastRowNum; rowInx++)
                {
                    IRow sheetRow = sheet.GetRow(rowInx);
                    if (sheetRow == null || sheetRow.Cells.Count <= 3) continue;//当excel行列数小于等于3则忽略
                    DataRow dataRow = dt.NewRow();
                    for (int colInx = 0; colInx < coluNum; colInx++)
                    {
                        ICell cell = sheetRow.GetCell(colInx);
                        if (cell == null) continue;
                        if (cell.CellType == CellType.NUMERIC)
                        {
                            //NPOI中数字和日期都是NUMERIC类型的，这里对其进行判断是否是日期类型
                            if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                            {
                                dataRow[colInx] = cell.DateCellValue;
                            }
                            else//其他数字类型
                            {
                                dataRow[colInx] = cell.NumericCellValue;
                            }
                        }
                        else
                        {
                            dataRow[colInx] = cell.ToString();
                        }
                    }
                    dt.Rows.Add(dataRow);
                }
                return dt;
            }
            catch (Exception ex)
            {
                //日志记录点
                throw ex;
            }
            finally
            {
                fileStream.Close();
                fileStream.Dispose();
            }
        }
    }

}