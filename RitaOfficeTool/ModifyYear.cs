using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{
    public class ModifyYear
    {
        public static List<string> oldStrList = new List<string>() { "年末余额", "本年发生数" };
        public static List<string> newStrList = new List<string>() { "年初余额", "上年发生数" };

        /// <summary>
        /// 是不是中文中待修改的年份的表格
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public bool NeedModifyYearTableForCn(Table table)
        {
            bool headLine1Error = false;
            var headerLine1 = WordTableUtil.GetRawRowHeader(table, 1, out headLine1Error);
            headerLine1 = headerLine1.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            Debug.WriteLine($"表头长度为:{headerLine1.Count}");
            return (headerLine1.Contains("年初余额") && headerLine1.Contains("年末余额"));
        }


        /// <summary>
        /// 是不是英文中待修改的年份的表格
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public bool NeedModifyYearTableForEn(Table table)
        {
            bool headLine1Error = false;
            var headerLine1 = WordTableUtil.GetRawRowHeader(table, 1, out headLine1Error);
            headerLine1 = headerLine1.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();

            return headerLine1.Contains("年初余额") && headerLine1.Contains("年末余额");
        }

        // /// <summary>
        // /// 横向表头是不是一行
        // /// </summary>
        // /// <param name="table"></param>
        // /// <returns></returns>
        // public bool RowHeaderIsSingleLine(Table table)
        // {
        //     bool headLine1Error = false;
        //     var headerLine1 = WordTableUtil.GetRawRowHeader(table, 1, out headLine1Error);
        //     // headLine1Error=true
        //     return !headLine1Error;
        // }



        /// <summary>
        /// 获取当前行中待拷贝的索引对。固定查找第一行，或者第二行是否满足要求
        /// </summary>
        /// <param name="rawHeaderList"></param>
        /// <returns>前一个int,旧位置,后一个int,新位置</returns>
        public Dictionary<int, int> TryToGetPairByLine(List<string> rawHeaderList)
        {
            Dictionary<int, int> pairLists = new Dictionary<int, int>();
            var targetOldColIndex = FindColIndexs(rawHeaderList, oldStrList);
            if (targetOldColIndex.Count == 0) return pairLists; // 没找到
            var targetNewColIndex = FindColIndexs(rawHeaderList, newStrList);

            for (int i = 0; i < targetOldColIndex.Count; i++)
            {
                Debug.WriteLine($"旧位置:[{targetOldColIndex[i]}],新位置:[{targetNewColIndex[i]}]");
                pairLists.Add(targetOldColIndex[i], targetNewColIndex[i]);
            }

            return pairLists;
        }

        public static List<int> FindColIndexs(List<string> list, List<string> targetList)
        {
            List<int> positions = new List<int>();

            // 遍历列表，查找目标字符串的位置
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < targetList.Count; j++)
                {
                    if (list[i].Contains(targetList[j]))
                    {
                        // +1,在word中表格行和列都是从1开始计数
                        positions.Add(i + 1); // 将匹配的索引添加到结果列表
                        break;
                    }
                }
            }

            return positions;
        }

        public bool TableIsNeedToCopy(Table table)
        {
            return ValidRow(table, 1) || ValidRow(table, 2);
        }

        public bool ValidRow(Table table, int rowIndex)
        {
            return GetPairByRowIndex(table, rowIndex).Count > 0;
        }

        private Dictionary<int, int> GetPairByRowIndex(Table table, int rowIndex)
        {
            bool headLine1Error = false;
            var headerLine1 = WordTableUtil.GetRawRowHeader(table, rowIndex, out headLine1Error);
            var toCopyPair = TryToGetPairByLine(headerLine1);
            return toCopyPair;
        }

        public Dictionary<int, int> ValidPair(Table table)
        {
            var pairByRowIndex1 = GetPairByRowIndex(table, 1);
            if (pairByRowIndex1.Count > 0) return pairByRowIndex1; // 第一行有值,就返回
            return GetPairByRowIndex(table, 2); // 第二行返回时,不关心有没有值
        }

        public void ReplaceColValue(Table table, Dictionary<int, int> pair)
        {
            int rowIndex = WordTableUtil.DataStartRow(table);

            foreach (KeyValuePair<int, int> kvp in pair)
            {
                for (int i = rowIndex; i <= table.Rows.Count; i++) // 行
                {
                    var oldRangText = WordTableUtil.CellText(table, i, kvp.Key);
                    // if (oldRangText.Equals(""))
                    // {
                    //     System.Windows.Forms.MessageBox.Show("该文档已经使用过该功能了");
                    //     return;
                    // }
                    //

                    table.Cell(i, kvp.Value).Range.Text = oldRangText;
                    table.Cell(i, kvp.Key).Range.Text = "";
                }
            }

            // 清空其他列
            ClearOtherCol(table, pair, rowIndex);
        }

        private static void ClearOtherCol(Table table, Dictionary<int, int> pair, int rowIndex)
        {
            var newColIndexList = pair.Values;
            var minCol = pair.Values.Min();
            for (int col = minCol; col <= table.Columns.Count; col++) //列
            {
                if (!newColIndexList.Contains(col))
                {
                    for (int i = rowIndex; i <= table.Rows.Count; i++) // 行
                    {
                        table.Cell(i, col).Range.Text = "";
                    }
                }
            }
        }

        public void ReplaceSubColValue(Table table)
        {
            var oldColStartIndex = 2;
            var gap = table.Columns.Count / 2; //3
            var newColStartIndex = gap + oldColStartIndex;
            int rowIndex = WordTableUtil.DataStartRow(table);

            // 待拷贝的列的次数
            for (int j = 0; j < gap; j++) //列
            {
                for (int i = rowIndex; i <= table.Rows.Count; i++) // 行
                {
                    var oldRangText = WordTableUtil.CellText(table, i, oldColStartIndex + j);
                    if (oldRangText.Equals(""))
                    {
                        System.Windows.Forms.MessageBox.Show("该文档已经使用过该功能了");
                        return;
                    }
                    table.Cell(i, newColStartIndex + j).Range.Text = oldRangText;
                    table.Cell(i, oldColStartIndex + j).Range.Text = "";
                }
            }
        }
    }
}