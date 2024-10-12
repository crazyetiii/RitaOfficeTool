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


        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        // 更新 Word 状态栏
        private void UpdateStatusBar(string message)
        {
            Application wordApp = Globals.ThisAddIn.Application;

            // 将信息设置到状态栏
            wordApp.StatusBar = message;
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前 Word 应用程序对象
            Application wordApp = Globals.ThisAddIn.Application;

            // 例如，获取当前活动文档
            Document activeDoc = wordApp.ActiveDocument;

            // 获取选中的文本
            Selection selection = wordApp.Selection;

            // 判断当前是否选中表格中的数据
            if (selection.Tables.Count < 0)
            {
                // 未选中表格，弹出提示
                System.Windows.Forms.MessageBox.Show("请先选中表格中的数据", "提示", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            // 选中了表格中的数据，获取第一个选中的表格
            Table selectedTable = selection.Tables[1];
            // 打印横向表头
            var msg = _tableSum.SumSelectedCellsData(selection, selectedTable);
            UpdateStatusBar(msg);
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

        public void InsertColumnInTable(Document document, int tableIndex, int columnIndex)
        {
            // 获取指定的表格
            Table table = document.Tables[tableIndex];

            // 检查列索引是否合法
            if (columnIndex < 1 || columnIndex > table.Columns.Count + 1)
            {
                System.Windows.Forms.MessageBox.Show("列索引不合法，必须在1到" + (table.Columns.Count + 1) + "之间。");
                return;
            }

            // 插入新列
            table.Columns.Add(table.Cell(1, columnIndex).Range);

            // 可以选择填充新列的内容
            for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
            {
                table.Cell(rowIndex, columnIndex).Range.Text = "新列数据"; // 或者根据需要填充数据
            }

            // 提示用户操作已完成
            System.Windows.Forms.MessageBox.Show("已在表格中插入新列。");
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取当前 Word 应用程序对象
            Application wordApp = Globals.ThisAddIn.Application;

            // 获取当前活动文档
            Document activeDoc = wordApp.ActiveDocument;

        }
    }
}