using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{

    class TableSum
    {
        
        // 打印选中的单元格数据
        public string SumSelectedCellsData(Word.Selection selection, Word.Table table)
        {
            var rowHeaders = WordTableUtil.GetValidRowHeader(table);
            var colHeaders = WordTableUtil.GetRawColHeader(table);

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
            return itemContent.Contains("本期减少") || itemContent.Contains("本年减少") || itemContent.Contains("减：");
        }
    }
}