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

       
    }
}