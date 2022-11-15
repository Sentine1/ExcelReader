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

                        IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalCollumns].Select(c => c.Text == null ? string.Empty : c.Text);
                        var list = row.ToList();
                        var list2 = ew.ToList();
                        for (int i = 0; i < list2.Count; i++)
                        {
                            excelTableData[rowIndex - 1, i] = list2[i] == string.Empty ? "*" : list2[i];
                        }
                    }
                    writeText();
                }
                else
                    throw new Exception("File not found");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
        }

        protected void writeText()
        {
            richTextBox1.Clear();
            for (int i = 0; i < totalRows; i++)
            {
                for (int j = 0; j < totalCollumns; j++)
                {
                    if (textBox1.Text != string.Empty )
                        if (excelTableData[i, 2] != userName ||
                        excelTableData[i, 3] != usserLogin)
                        {
                            j += totalCollumns;
                            continue;
                        }
                    richTextBox1.Text += excelTableData[i, j] + " ";
                }
                richTextBox1.Text += "\n";
            }
        }

        private void тестToolStripMenuItem_Click(object sender, EventArgs e)
        {
            userName = textBox1.Text;
            writeText();
        }

        private void тестToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            usserLogin = textBox1.Text;
            writeText();
        }
    }
}
