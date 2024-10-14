using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace RitaOfficeTool
{
    public partial class MyRibbon
    {
        private TableSum _tableSum = new TableSum();
        private ModifyYear _modifyYear = new ModifyYear();

        /// <summary>
        /// 清除所有文档表格中数字中的空白
        /// </summary>
        /// <param name="tables"></param>
        private static void ClearAllTableData(Tables tables)
        {
            string pattern = @"(\S)\s+(\r\a)";
            for (int k = 1; k < tables.Count; k++)
            {
                var table = tables[k];
                var rowsCount = table.Rows.Count;
                var columnsCount = table.Columns.Count;
                Debug.WriteLine($"当前表格【{k}】:行数:{rowsCount},列数:{columnsCount}");
                // 行
                for (int rowIndex = 1; rowIndex <= rowsCount; rowIndex++)
                {
                    // 列
                    for (int col = 1; col < columnsCount; col++)
                    {
                        try
                        {
                            // 获取当前单元格
                            Cell cell = table.Cell(rowIndex, col);
                            // 获取单元格的内容，并移除末尾的段落标记和其他控制字符
                            var rawText = cell.Range.Text;
                            Match match = Regex.Match(rawText, pattern);
                            if (!match.Success) continue;

                            string result = Regex.Replace(rawText, pattern, "$1$2");
                            cell.Range.Text = result;
                        }
                        catch (System.Exception ex)
                        {
                            Debug.WriteLine($"error");
                        }
                    }
                }
            }
        }

        // 更新 Word 状态栏
        private void UpdateStatusBar(string message)
        {
            Application wordApp = Globals.ThisAddIn.Application;
            // 将信息设置到状态栏
            wordApp.StatusBar = message;
        }

        private static bool SelectionIsTableData(out Selection selection)
        {
            // 获取当前 Word 应用程序对象
            Application wordApp = Globals.ThisAddIn.Application;
            // 获取选中的文本
            selection = wordApp.Selection;

            // 判断当前是否选中表格中的数据
            if (selection.Tables.Count == 0)
            {
                // 未选中表格，弹出提示
                System.Windows.Forms.MessageBox.Show("请先选中表格中的数据", "提示", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前 Word 应用程序对象
            Application wordApp = Globals.ThisAddIn.Application;

            // 获取当前活动文档
            Document activeDoc = wordApp.ActiveDocument;

            // 获取文档中的所有表格
            Tables tables = activeDoc.Tables;
            ClearAllTableData(tables);
            activeDoc.Save(); // 直接保存
            UpdateStatusBar($"清理完成!");
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (!SelectionIsTableData(out var selection)) return;

            // 选中了表格中的数据，获取第一个选中的表格
            Table selectedTable = selection.Tables[1];
            // 打印横向表头
            var msg = _tableSum.SumSelectedCellsData(selection, selectedTable);
            UpdateStatusBar(msg);
        }


        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前 Word 应用程序对象
            Application wordApp = Globals.ThisAddIn.Application;

            // 获取当前活动文档
            Document activeDoc = wordApp.ActiveDocument;
            Tables tables = activeDoc.Tables;

            // 1.中文文档
            foreach (Table table in tables)
            {
                if (_modifyYear.ValidRow(table, 1)) // 关键字在第1行
                {
                    if (WordTableUtil.RowHeaderIsSingleLine(table)) // 表头只有1行。对称和非对称都可用.
                    {
                        var validPair = _modifyYear.ValidPair(table);
                        _modifyYear.ReplaceColValue(table, validPair);
                    }
                    else // 表头非1行,这里指两行 2024年10月14日08:14:01
                    {
                        _modifyYear.ReplaceSubColValue(table);
                    }
                }
                else if (_modifyYear.ValidRow(table, 2)) // 关键字在第2行
                {
                    var validPair = _modifyYear.ValidPair(table);
                    _modifyYear.ReplaceColValue(table, validPair);
                }
                else // 没有找到关键字
                {
                    continue;
                }
            }

            activeDoc.Save(); // 直接保存
            UpdateStatusBar($"添加年份完成!");
        }


        private void button4_Click_1(object sender, RibbonControlEventArgs e)
        {
            if (!SelectionIsTableData(out var selection)) return;

            // 选中了表格中的数据，获取第一个选中的表格
            Table table = selection.Tables[1];
            if (selection.Columns.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show("只能选择1列", "提示", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            if (selection.Cells.Count == 1) // 点击排序列的任意单元格
            {
                Cell selectedCell = selection.Cells[1];
                // 获取该单元格所在列的索引
                int toOrderColIndex = selectedCell.ColumnIndex;
                OrderTable.OrderAllData(table, toOrderColIndex);
            }
            else // 只排序选中的单元格
            {
                var cellRangeInfo = WordTableUtil.GetCellRangeInfo(selection);
                OrderTable.OrderPartData(table, cellRangeInfo.MinColumn, cellRangeInfo.MinRow, cellRangeInfo.MaxRow);
            }

            // 打印横向表头
            var msg = "排序完成!";
            UpdateStatusBar(msg);
        }
    }
}