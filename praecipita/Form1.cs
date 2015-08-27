using Microsoft.Office.Interop.Excel;
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

namespace praecipita
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = "D:\\_GIT\\Projeto Andorinhas (GPD3)\\Dados CD";
            if(folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int varRow = 10;
            int timeRow = 11;
            int dateCol = 1;
            int startRow = 12;
            int startCol = 2;
            int maxRow = 0;
            int maxCols = 0;
            int rangeStartRow = 1;
            int rangeStartCol = 1;
            int rangeMaxRow;
            int rangeMaxCol;
            int rowsIndex = 0;
            int colsIndex = 0;
            string currLine = "";
            string currValueStr = "";
            DateTime currDate;
            string currTime = "";
            string stationName = "";
            string variableName = "";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            DirectoryInfo rootDir = new DirectoryInfo(textBox1.Text);
            DirectoryInfo[] folders = rootDir.GetDirectories();
            foreach(DirectoryInfo currFolder in folders)
            {
                foreach (FileInfo currFile in currFolder.GetFiles())
                {
                    if(currFile.Name.IndexOf(".xls") > 0)
                    {
                        stationName = currFile.Name.Substring(0, currFile.Name.Length - 4);
                        Workbook currWB = excel.Workbooks.Open(currFile.FullName);
                        Worksheet currWS = currWB.Sheets[1];

                        // The following 2 lines exclude the null cells within used range
                        
                        currWS.Rows.ClearFormats();
                        currWS.Columns.ClearFormats();

                        maxRow = currWS.UsedRange.Rows.Count;
                        maxCols = currWS.UsedRange.Columns.Count;

                        Range usedRange = currWS.get_Range(currWS.Cells[varRow,dateCol],currWS.Cells[maxRow,maxCols]);
                        object[,] valueArray = (object[,])usedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                        rangeMaxRow = maxRow - startRow + 1;
                        rangeMaxCol = maxCols - dateCol + 1;
                        for (rowsIndex = rangeStartRow; rowsIndex <= rangeMaxRow; rowsIndex++)
                        {
                            currDate = currWS.Cells[rowsIndex, dateCol].Value;
                            for (colsIndex = startCol; colsIndex <= maxCols; colsIndex++)
                            {
                                if (currWS.Cells[rowsIndex, colsIndex].Value.ToString() == "NULL")
                                    currValueStr = currWS.Cells[rowsIndex, colsIndex].Value;
                                else
                                    currValueStr = currWS.Cells[rowsIndex, colsIndex].Value.ToString().Replace(",", ".");
                                currTime = currWS.Cells[timeRow, colsIndex].Value.ToString("0000");
                                currTime = currTime.Substring(0, 2) + ":" + currTime.Substring(2, 2);
                                variableName = currWS.Cells[varRow, colsIndex].Value.Replace(" ", "_");
                                currLine = stationName + ";" + currDate.ToString("dd/MM/yyyy") + " " + currTime + ";" + variableName + ";" + currValueStr;
                                //textBox2.Text = currLine;
                                textBox2.Text = currDate.ToString("dd/MM/yyyy");
                            }
                        }
                        currWB.Close();
                    }
                }
            }
        }

    }
}
