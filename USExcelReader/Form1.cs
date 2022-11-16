using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace USExcelReader
{
    public partial class Form1 : Form
    {
        private String[,] excelTableData;
        private int totalRows = 0;
        private int totalCollumns = 0;
        private string userName = null;
        private string usserLogin = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();

                if (res == DialogResult.OK)
                {
                    ExcelPackage excelFile = new ExcelPackage(new FileInfo(openFileDialog1.FileName));

                    ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[0];

                    totalRows = worksheet.Dimension.End.Row;
                    totalCollumns = worksheet.Dimension.End.Column;

                    excelTableData = new string[totalRows, totalCollumns];

                    for (int rowIndex = 1; rowIndex < totalRows + 1; rowIndex++)
                    {
                        var ew = new List<string>();
                        for (int collumTest = 1; collumTest < totalCollumns + 1; collumTest++)
                        {
                            ew.Add(worksheet.Cells[rowIndex, collumTest].Text);
                        }

                        //IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalCollumns].Select(c => c.Text == null ? string.Empty : c.Text);
                        //var list = row.ToList();
                        var list2 = ew.ToList();
                        for (int i = 0; i < list2.Count; i++)
                        {
                            excelTableData[rowIndex - 1, i] = list2[i] == string.Empty ? "*" : list2[i];
                        }
                    }
                    //writeText();
                }
                else
                    throw new Exception("File not found");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
        }

        protected void writeText(string filter = "", int position = -1)
        {
            richTextBox1.Clear();
            richTextBox1.Text += excelTableData[0, 1] + "\t\t\t\t" + excelTableData[0, 2] + "\t\t" + excelTableData[0, 18] + "\t\t" + excelTableData[0, 19] + "\n";
            var Dictionary = new Dictionary<string,int>();
            for (int i = 1; i < totalRows; i++)
            {
                if (filter != string.Empty && position != -1)
                    if (!excelTableData[i, position].Contains(filter))
                    {
                        continue;
                    }
                    else
                    {
                        if (Dictionary.ContainsKey(excelTableData[i, 1] + excelTableData[i, 2] + excelTableData[i, 18]))
                            Dictionary[excelTableData[i, 1] + excelTableData[i, 2] + excelTableData[i, 18]]++;
                        else
                        {
                            richTextBox1.Text += excelTableData[i, 1] + "\t" + excelTableData[i, 2] + "\t" + excelTableData[i, 18] + "\t" + excelTableData[i, 19] + "\n";
                            Dictionary.Add((excelTableData[i, 1] + excelTableData[i, 2] + excelTableData[i, 18]), 1);
                        }
                    }
            }
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            userName = textBox1.Text;
            writeText(userName, 1);
        }

        private void тестToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            usserLogin = textBox1.Text;
            writeText(usserLogin, 2);
        }
    }
}
