using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{
    public class WordTableUtil
    {
        /// <summary>
        /// 获取原生的横向表头
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowIndex">第几行</param>
        /// <param name="error">为true代表,有合并单元格,还需要判断下一行</param>
        /// <returns></returns>
        public static List<string> GetRawRowHeader(Table table, int rowIndex, out bool error)
        {
            bool getItemError = false;
            // 创建一个列表来存储表头
            List<string> headerLine = new List<string>();

            // 遍历表格的第一行
            for (int colIndex = 1; colIndex <= table.Columns.Count; colIndex++)
            {
                try
                {
                    // 获取当前单元格
                    Cell cell = table.Cell(rowIndex, colIndex);

                    // 获取单元格的内容，并移除末尾的段落标记和其他控制字符
                    string cellText = cell.Range.Text.TrimEnd('\r', '\a');

                    headerLine.Add(cellText);
                }
                catch (System.Exception ex)
                {
                    headerLine.Add("");
                    getItemError = true;
                }
            }

            error = getItemError;
            return headerLine;
        }

        /// <summary>
        /// 获取纵向表头
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static List<string> GetRawColHeader(Table table)
        {
            // 创建一个列表来存储表头
            List<string> headerLine = new List<string>();

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                try
                {
                    // 获取当前单元格
                    Cell cell = table.Cell(row, 1);

                    // 获取单元格的内容，并移除末尾的段落标记和其他控制字符
                    string cellText = cell.Range.Text.TrimEnd('\r', '\a');

                    headerLine.Add(cellText);
                }
                catch (System.Exception ex)
                {
                    headerLine.Add("");
                }
            }

            return headerLine;
        }

        /// <summary>
        /// 获取header中第一个有效值的索引,这里返回的是table中的索引.从1开始
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        private static int GetValidIndex(List<string> header)
        {
            for (int i = 0; i < header.Count; i++)
            {
                if (header[i] != "")
                {
                    return i + 1;
                }
            }

            return -1;
        }

        /// <summary>
        /// 获取横向有效表头
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static List<string> GetValidRowHeader(Table table)
        {
            bool headLine1Error = false;
            var headerLine1 = GetRawRowHeader(table, 1, out headLine1Error);
            if (!headLine1Error) return headerLine1;

            bool headLine2Error = false;
            var headerLine2 = GetRawRowHeader(table, 2, out headLine2Error);
            var validIndex = GetValidIndex(headerLine2);
            headerLine1.InsertRange(validIndex, headerLine2.Where(item => !string.IsNullOrWhiteSpace(item)).ToList());
            headerLine1.RemoveAt(validIndex - 1);
            headerLine1 = headerLine1.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            return headerLine1;
        }

        /// <summary>
        /// 横向表头是不是一行
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static bool RowHeaderIsSingleLine(Table table)
        {
            bool headLine1Error = false;
            var headerLine1 = WordTableUtil.GetRawRowHeader(table, 1, out headLine1Error);
            return !headLine1Error;
        }

        public static string CellText(Table table, int row, int col)
        {
            try
            {
                var rangeText = table.Cell(row, col).Range.Text;
                return rangeText.Replace("\r\a", "");
            }
            catch (Exception e)
            {
                return "";
            }
        }

        public static void SetCellText(Table table, int row, int col, string result)
        {
            table.Cell(row, col).Range.Text = result;
        }

        /// <summary>
        /// 横向数据从第几行开始
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static int DataStartRow(Table table)
        {
            return RowHeaderIsSingleLine(table) ? 2 : 3;
        }

        // 定义一个类，用于存储单元格范围信息
        public class CellRangeInfo
        {
            public int MinRow { get; set; }
            public int MaxRow { get; set; }
            public int MinColumn { get; set; }
            public int MaxColumn { get; set; }
        }

        // 获取选中单元格的范围信息
        public static CellRangeInfo GetCellRangeInfo(Selection selection)
        {
            Cells selectedCells = selection.Cells;

            // 初始化最小值为最大值
            int minRow = int.MaxValue;
            int maxRow = int.MinValue;
            int minColumn = int.MaxValue;
            int maxColumn = int.MinValue;

            // 遍历所有选中的单元格
            foreach (Cell cell in selectedCells)
            {
                int rowIndex = cell.RowIndex;
                int columnIndex = cell.ColumnIndex;

                // 更新最小和最大行列索引
                if (rowIndex < minRow) minRow = rowIndex;
                if (rowIndex > maxRow) maxRow = rowIndex;
                if (columnIndex < minColumn) minColumn = columnIndex;
                if (columnIndex > maxColumn) maxColumn = columnIndex;
            }

            // 创建并返回 CellRangeInfo 对象
            return new CellRangeInfo
            {
                MinRow = minRow,
                MaxRow = maxRow,
                MinColumn = minColumn,
                MaxColumn = maxColumn
            };
        }

        
    }
}