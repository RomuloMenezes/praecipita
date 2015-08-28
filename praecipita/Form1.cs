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
            bool wbOpen = false;
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
                        wbOpen = true;
                        Worksheet currWS = currWB.Sheets[1];
                        maxRow = currWS.UsedRange.Rows.Count;
                        maxCol = currWS.UsedRange.Columns.Count;

                        Range usedRange = currWS.UsedRange;
                        try
                        {
                            for (rowsIndex = startRow; rowsIndex <= maxRow; rowsIndex++)
                            {
                                try
                                {
                                    currDate = usedRange.Cells[rowsIndex, dateCol].Value;
                                    for (colsIndex = startCol; colsIndex <= maxCol; colsIndex++)
                                    {
                                        if (usedRange.Cells[rowsIndex, colsIndex].Value.ToString() == "NULL")
                                            currValueStr = usedRange.Cells[rowsIndex, colsIndex].Value;
                                        else
                                            currValueStr = usedRange.Cells[rowsIndex, colsIndex].Value.ToString().Replace(",", ".");
                                        currTime = usedRange.Cells[timeRow, colsIndex].Value.ToString("0000");
                                        currTime = currTime.Substring(0, 2) + ":" + currTime.Substring(2, 2);
                                        variableName = usedRange.Cells[varRow, colsIndex].Value.Replace(" ", "_");
                                        currLine = stationName + ";" + currDate.ToString("dd/MM/yyyy") + " " + currTime + ";" + variableName + ";" + currValueStr;
                                        //textBox2.Text = currLine;
                                        textBox2.Text = currDate.ToString("dd/MM/yyyy");
                                    }
                                }
                                catch(Exception)
                                {

                                }
                            }
                            currWB.Close();
                        }
                        catch(Exception) // There are empty cells in the used rows
                        {
                            textBox2.Text = stationName + " processed";
                            currWB.Close();
                            wbOpen = false;
                        }
                    }
                }
            }
        }

    }
}
