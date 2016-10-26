using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace NPOI.SS.UserModel
{
    public static class NPOIRowExtend
    {
        public static void CreateCells(this IRow row,object data)
        {
            Type t =data.GetType();
            int i=0;
            foreach (var Propertie in t.GetProperties())
            {
                var cell=row.CreateCell(i++);
                cell.SetCellValue(Convert.ChangeType(Propertie.GetValue(data),Propertie.PropertyType).ToString());
            }
        }

        public static void MoveCell(this ISheet sheet, int curentIndex, int moveIndex)
        {
            var row = sheet.GetRow(curentIndex);//需要移植的行
            var moveRow = sheet.GetRow(moveIndex);
            if (moveRow == null)
            {
               moveRow= sheet.CreateRow(moveIndex);
            }
            for (int i = 0; i <row.LastCellNum; i++)
            {
              var movecell=  moveRow.GetCell(i)??moveRow.CreateCell(i);
              var cell=  row.GetCell(i);
              movecell.SetCellValue(cell.StringCellValue);
            }
        }
    }
}