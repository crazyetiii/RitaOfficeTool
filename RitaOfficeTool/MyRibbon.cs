using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
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
            if (selection.Tables.Count > 0)
            {
                // 选中了表格中的数据，获取第一个选中的表格
                Table selectedTable = selection.Tables[1];
                // 打印横向表头
                var msg = _tableSum.SumSelectedCellsData(selection, selectedTable);
                UpdateStatusBar(msg);
            }
            else
            {
                // 未选中表格，弹出提示
                System.Windows.Forms.MessageBox.Show("请先选中表格中的数据", "提示", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }
    }
}