using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{
    class TableSum
    {
        /// <summary>
        /// 获取横向有效表头
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        private static List<string> GetValidRowHeader(Word.Table table)
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
        /// 获取原生的横向表头
        /// </summary>
        /// <param name="table"></param>
        /// <param name="rowIndex">第几行</param>
        /// <param name="error">为true代表,有合并单元格,还需要判断下一行</param>
        /// <returns></returns>
        private static List<string> GetRawRowHeader(Word.Table table, int rowIndex, out bool error)
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
                    Word.Cell cell = table.Cell(rowIndex, colIndex);

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

        private static List<string> GetRawColHeader(Word.Table table)
        {
            // 创建一个列表来存储表头
            List<string> headerLine = new List<string>();

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                try
                {
                    // 获取当前单元格
                    Word.Cell cell = table.Cell(row, 1);

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

        // 打印选中的单元格数据
        public string SumSelectedCellsData(Word.Selection selection, Word.Table table)
        {
            var rowHeaders = GetValidRowHeader(table);
            var colHeaders = GetRawColHeader(table);

            double total = 0.0;
            // 获取当前选中的单元格范围
            foreach (Word.Cell cell in selection.Cells)
            {
                // 获取单元格的范围
                Word.Range cellRange = cell.Range;

                // 获取单元格的内容，并移除末尾的段落标记和其他控制字符
                string cellText = cellRange.Text.TrimEnd('\r', '\a');

                if (double.TryParse(cellText, out double cellValue))
                {
                    bool rowHeaderContainsReduction = ContainsReductionTerms(cell.ColumnIndex, rowHeaders);
                    bool colHeaderContainsReduction = ContainsReductionTerms(cell.RowIndex, colHeaders);

                    // 如果行或列满足条件，取负值
                    if (rowHeaderContainsReduction || colHeaderContainsReduction)
                    {
                        cellValue = -cellValue;
                    }

                    Debug.WriteLine($"单元格的值:{cellValue:N2}");

                    total += cellValue;
                }
            }

            return $"合计: {total:N2}";
        }

        private static bool ContainsReductionTerms(int index, List<string> headers)
        {
            if (index < 1 || index > headers.Count) return false;

            string itemContent = headers[index - 1];
            return itemContent.Contains("本期减少") || itemContent.Contains("本年减少") || itemContent.Contains("减");
        }
    }
}