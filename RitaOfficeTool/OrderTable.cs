using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

namespace RitaOfficeTool
{
    public class OrderTable
    {
        private static void SetTableData(Table table, List<List<string>> tableData)
        {
            int rowIndex = 0;
            int colIndex = 0;

            var dataStartRow = WordTableUtil.DataStartRow(table);
            for (int i = dataStartRow; i < table.Rows.Count; i++)
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    WordTableUtil.SetCellText(table, i, j, tableData[rowIndex][colIndex]);
                    colIndex++;
                }

                rowIndex++;
                colIndex = 0;
            }
        }

        public static void OrderAllData(Table table, int col)
        {
            var orderData = GetOrderedDataByColIndex(table, col);
            SetTableData(table, orderData);
        }

        /// <summary>
        /// 获取表格中的所有数据,不包含最后一行
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static List<List<string>> GetTableData(Table table)
        {
            ValidDataPos(table, out var startRowIndex, out var startColIndex);

            List<List<string>> result = new List<List<string>>();

            for (int i = startRowIndex; i < table.Rows.Count; i++) // 行
            {
                List<string> curLineData = new List<string>();

                for (int j = 1; j <= table.Columns.Count; j++) // 列
                {
                    try
                    {
                        var rowText = WordTableUtil.CellText(table, i, j);
                        curLineData.Add(rowText);
                    }
                    catch (Exception e)
                    {
                        curLineData.Add("");
                    }
                }

                result.Add(curLineData);
            }

            return result;
        }

        private static void SetPartTableData(Table table, List<List<string>> tablePartData, int startRow, int endRow)
        {
            int rowIndex = 0;
            int colIndex = 0;

            for (int i = startRow; i <= endRow; i++) // 行
            {
                for (int j = 1; j <= table.Columns.Count; j++)
                {
                    WordTableUtil.SetCellText(table, i, j, tablePartData[rowIndex][colIndex]);
                    colIndex++;
                }

                rowIndex++;
                colIndex = 0;
            }
        }

        public static void OrderPartData(Table table, int col, int minRow, int maxRow)
        {
            var toOrderData = GetPartTableData(table, minRow, maxRow);
            var orderData = OrderData(toOrderData, col-1);
            SetPartTableData(table, orderData, minRow, maxRow);
        }

        /// <summary>
        /// 获取部分数据
        /// </summary>
        /// <param name="table"></param>
        /// <param name="minRow">开始行</param>
        /// <param name="maxRow">结束行</param>
        /// <returns></returns>
        public static List<List<string>> GetPartTableData(Table table, int minRow, int maxRow)
        {
            var calMinRow = Math.Max(WordTableUtil.DataStartRow(table), minRow);

            List<List<string>> result = new List<List<string>>();

            for (int i = calMinRow; i <= maxRow; i++) // 行
            {
                List<string> curLineData = new List<string>();

                for (int j = 1; j <= table.Columns.Count; j++) // 列
                {
                    try
                    {
                        var rowText = WordTableUtil.CellText(table, i, j);
                        curLineData.Add(rowText);
                    }
                    catch (Exception e)
                    {
                        curLineData.Add("");
                    }
                }

                result.Add(curLineData);
            }

            return result;
        }

        private static void ValidDataPos(Table table, out int startRowIndex, out int startColIndex)
        {
            startColIndex = 2;
            startRowIndex = WordTableUtil.DataStartRow(table);
        }

        /// <summary>
        /// 将table表按照colIndex列排序
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colIndex">选中的某列的索引.大于1</param>
        /// <returns></returns>
        public static List<List<string>> GetOrderedDataByColIndex(Table table, int colIndex)
        {
            var oldTableData = GetTableData(table);
            return OrderData(oldTableData, colIndex - 1);
        }

        private static List<List<string>> OrderData(List<List<string>> oldTableData, int columnToSortBy)
        {
            oldTableData.Sort((list1, list2) =>
            {
                double value1, value2;

                // 尝试解析 list1 和 list2 的值，如果解析失败则设置默认值为0
                bool isParsed1 = double.TryParse(list1[columnToSortBy], out value1);
                bool isParsed2 = double.TryParse(list2[columnToSortBy], out value2);
                return value2.CompareTo(value1);
            });

            // 打印排序后的结果
            foreach (var list in oldTableData)
            {
                Debug.WriteLine(string.Join(", ", list));
            }

            return oldTableData;
        }


        /// <summary>
        /// 获取某列的纯数据部分
        /// </summary>
        /// <param name="table"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        public static List<string> GetColData(Table table, int colIndex)
        {
            var startRow = WordTableUtil.DataStartRow(table);
            List<string> result = new List<string>();

            for (int i = startRow; i <= table.Rows.Count; i++)
            {
                var cellText = WordTableUtil.CellText(table, startRow, colIndex);
                result.Add(cellText);
            }

            return result;
        }
    }
}