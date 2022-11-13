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
                        IEnumerable<string> row = worksheet.Cells[rowIndex, 1, rowIndex, totalCollumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        var list = row.ToList();
                        for (int i = 0; i < list.Count; i++)
                        {
                            excelTableData[rowIndex - 1, i] = list[i];
                        }
                    }

                    for (int i = 0; i < totalRows; i++)
                    {
                        for (int j = 0; j < totalCollumns; j++)
                        {
                            richTextBox1.Text += excelTableData[i, j] + " ";
                        }
                        richTextBox1.Text += "\n";
                    }
                }
                else
                    throw new Exception("File not found");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }  
        }
    }
}
