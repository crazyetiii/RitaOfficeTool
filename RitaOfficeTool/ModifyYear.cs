using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{
    public class ModifyYear
    {
        public static List<string> oldStrList = new List<string>(Rita.Default.old_sub_str.Split('|'));
        public static List<string> newStrList = new List<string>(Rita.Default.new_sub_str.Split('|'));


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
            if (targetNewColIndex.Count == 0) return pairLists; // 没找到

            for (int i = 0; i < targetOldColIndex.Count; i++)
            {
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

        public bool IsEmpty(Table table, int col,int startRow)
        {
            for (int i = startRow; i < table.Rows.Count; i++) // 最后一行不算,合计
            {
                var oldRangText = WordTableUtil.CellText(table, i, col);
                if (!oldRangText.Equals("")) // 有值
                {
                    return true;
                }
            }
            return false;
        }


        public void ReplaceColValue(Table table, Dictionary<int, int> pair)
        {
            int rowIndex = WordTableUtil.DataStartRow(table);
            bool clear = false;

            foreach (KeyValuePair<int, int> kvp in pair) // 旧列,新列
            {
                // 从rowIndex开始,检查该列是否为空,为空代表已经使用过了
                for (int i = rowIndex; i <= table.Rows.Count; i++) // 行
                {
                    try
                    {
                        var oldRangText = WordTableUtil.CellText(table, i, kvp.Key);
                        Debug.WriteLine(oldRangText);
                        table.Cell(i, kvp.Value).Range.Text = oldRangText;
                        table.Cell(i, kvp.Key).Range.Text = "";
                    }
                    catch (System.Exception)
                    {
                        continue;
                    }
                }
                clear = true;
            }

            // 清空其他列
            if (clear)
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
                        try
                        {
                            table.Cell(i, col).Range.Text = "";
                        }
                        catch (System.Exception)
                        {
                            continue;
                        }
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
                    try
                    {
                        var oldRangText = WordTableUtil.CellText(table, i, oldColStartIndex + j);
                        table.Cell(i, newColStartIndex + j).Range.Text = oldRangText;
                        table.Cell(i, oldColStartIndex + j).Range.Text = "";
                    }
                    catch (System.Exception)
                    {
                        continue;
                    }
                }

            }
        }
    }
}