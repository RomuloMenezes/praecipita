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
            int maxCol = 0;
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
            int nbOfFolders = folders.Count();
            int nbOfFiles = 0;
            int folderCount = 0;
            int fileCount = 0;
            foreach(DirectoryInfo currFolder in folders)
            {
                nbOfFiles = currFolder.GetFiles().Count();
                folderCount += 1;
                foreach (FileInfo currFile in currFolder.GetFiles())
                {
                    if(currFile.Name.IndexOf(".xls") > 0)
                    {
                        fileCount += 1;
                        stationName = currFile.Name.Substring(1, currFile.Name.Length - 5);
                        if (stationName.Substring(stationName.Length - 1, 1) == "_")
                            stationName = stationName.Substring(0, stationName.Length - 1);
                        if (stationName.IndexOf("_2015") > -1)
                            stationName = stationName.Replace("_2015", "");
                        Workbook currWB = excel.Workbooks.Open(currFile.FullName);
                        Worksheet currWS = currWB.Sheets[1];
                        Range usedRange = currWS.UsedRange;
                        var rangeArray = usedRange.Value;
                        currWS.Rows.ClearFormats();
                        currWS.Columns.ClearFormats();
                        maxRow = currWS.UsedRange.Rows.Count;
                        maxCol = currWS.UsedRange.Columns.Count;
                        
                        if (checkBox1.Checked)
                        {
                            textBox2.Text = stationName + Environment.NewLine;
                            textBox2.Text += "File " + fileCount.ToString() + " of " + nbOfFiles.ToString() + " in folder " + folderCount.ToString() + " of " + nbOfFolders.ToString();
                            textBox2.Refresh();
                        }
                        for (rowsIndex = startRow; rowsIndex <= maxRow; rowsIndex++)
                        {
                            try
                            {
                                currDate = rangeArray[rowsIndex, dateCol];
                                for (colsIndex = startCol; colsIndex <= maxCol; colsIndex++)
                                {
                                    try
                                    {
                                        if (rangeArray[rowsIndex, colsIndex].ToString() == "NULL")
                                            currValueStr = rangeArray[rowsIndex, colsIndex];
                                        else
                                            currValueStr = rangeArray[rowsIndex, colsIndex].ToString().Replace(",", ".");
                                        currTime = rangeArray[timeRow, colsIndex].ToString("0000");
                                        currTime = currTime.Substring(0, 2) + ":" + currTime.Substring(2, 2);
                                        variableName = rangeArray[varRow, colsIndex].Replace(" ", "_");
                                        currLine = stationName + ";" + currDate.ToString("dd/MM/yyyy") + " " + currTime + ";" + variableName + ";" + currValueStr;
                                    }
                                    catch // There are empty cols in used range
                                    {
                                        colsIndex = maxCol;
                                    }
                                }
                            }
                            catch // There are empty rows in used range
                            {
                                rowsIndex = maxRow;
                            }
                        }
                        currWB.Close();
                        fileCount = 0;
                    }
                }
            }
        }

    }
}
