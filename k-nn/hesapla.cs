using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace k_nn
{
    class hesapla
    {
        public void excelAll()
        { 
        
        }
        public int oklit(string[,] training, string[,] test, int komsu, int trainingSatirSay, int testSatirSay)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            Range myRange;
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 0];
            myRange.Value2 = "STG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 1];
            myRange.Value2 = "SCG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 2];
            myRange.Value2 = "STR";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 3];
            myRange.Value2 = "LPR";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 4];
            myRange.Value2 = "PEG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 5];
            myRange.Value2 = "UNS";
            StartRow++;
            double uzaklik = 0;
            int dogru = 0;
            int[] tut = new int[4];
            string[,] sonucMatrisi = new string[trainingSatirSay, 7];
            for (int i = 0; i < testSatirSay; i++)
            {
                for (int j = 0; j < trainingSatirSay; j++)
                {
                    for (int k = 0; k < 5; k++)
                    {
                        uzaklik += Math.Pow((Convert.ToDouble(training[j, k]) - Convert.ToDouble(test[i, k])), 2);
                        sonucMatrisi[j,k] = training[j, k];
                    }
                    uzaklik = Math.Sqrt(uzaklik);
                    sonucMatrisi[j, 6] = Convert.ToString(uzaklik);
                    sonucMatrisi[j, 5] = training[j, 5];
                    uzaklik = 0;
                }
                sonucMatrisi = sirala(sonucMatrisi, trainingSatirSay);
                for (int j = 0; j < komsu; j++)
                {
                    if (sonucMatrisi[j, 5] == "very_low")
                        tut[0]++;
                    else if (sonucMatrisi[j, 5] == "Low")
                        tut[1]++;
                    else if (sonucMatrisi[j, 5] == "Middle")
                        tut[2]++;
                    else
                        tut[3]++;
                }
                int itut = 0;
                for (int j = 0; j < 4; j++)
                    for (int k = 0; k < 4; k++)
                        if (tut[j] > tut[k])
                            itut = j;
                string sonucTut;
                if (itut == 0)
                    sonucTut = "very_low";
                else if (itut == 1)
                    sonucTut = "Low";
                else if (itut == 2)
                    sonucTut = "Middle";
                else
                    sonucTut = "High";
                if (sonucTut == test[i, 5])
                    dogru++;

                for (int j = 0; j < 5; j++)
                {
                    myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = test[i, j];
                    myRange.Select();
                }
                myRange = (Range)sheet1.Cells[StartRow + i, StartCol + 5];
                myRange.Value2 = sonucTut;
                myRange.Select();
                for (int j = 0; j < 4; j++)
                    tut[j] = 0;
            }
            

            return dogru;
        }
        public string[,] sirala(string[,] training, int satirSayisi)
        {
            string tut = "";
            for (int i = 0; i < satirSayisi; i++)
                for (int j = i; j < satirSayisi; j++)
                    if (Convert.ToDouble(training[j, 6]) < Convert.ToDouble(training[i, 6]))
                        for (int k = 0; k < 7; k++)
                        {
                            tut = training[j, k];
                            training[j, k] = training[i, k];
                            training[i, k] = tut;
                        }
            return training;
        }
        public int manhattan(string[,] training, string[,] test, int komsu, int trainingSatirSay, int testSatirSay)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            Range myRange;
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 0];
            myRange.Value2 = "STG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 1];
            myRange.Value2 = "SCG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 2];
            myRange.Value2 = "STR";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 3];
            myRange.Value2 = "LPR";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 4];
            myRange.Value2 = "PEG";
            myRange = (Range)sheet1.Cells[StartRow, StartCol + 5];
            myRange.Value2 = "UNS";
            StartRow++;
            double uzaklik = 0;
            int dogru = 0;
            int[] tut = new int[4];
            string[,] sonucMatrisi = new string[trainingSatirSay, 7];
            for (int i = 0; i < testSatirSay; i++)
            {
                for (int j = 0; j < trainingSatirSay; j++)
                {
                    for (int k = 0; k < 5; k++)
                    {
                        if (Convert.ToDouble(training[j, k]) - Convert.ToDouble(test[i, k]) > 0)
                            uzaklik += Convert.ToDouble(training[j, k]) - Convert.ToDouble(test[i, k]);
                        else
                            uzaklik += Convert.ToDouble(test[i, k]) - Convert.ToDouble(training[j, k]);
                        sonucMatrisi[j, k] = training[j, k];
                    }
                    sonucMatrisi[j, 6] = Convert.ToString(uzaklik);
                    sonucMatrisi[j, 5] = training[j, 5];
                    uzaklik = 0;
                }
                sonucMatrisi = sirala(sonucMatrisi, trainingSatirSay);
                for (int j = 0; j < komsu; j++)
                {
                    if (sonucMatrisi[j, 5] == "very_low")
                        tut[0]++;
                    else if (sonucMatrisi[j, 5] == "Low")
                        tut[1]++;
                    else if (sonucMatrisi[j, 5] == "Middle")
                        tut[2]++;
                    else
                        tut[3]++;
                }
                int itut = 0;
                for (int j = 0; j < 4; j++)
                    for (int k = 0; k < 4; k++)
                        if (tut[j] > tut[k])
                            itut = j;
                string sonucTut;
                if (itut == 0)
                    sonucTut = "very_low";
                else if (itut == 1)
                    sonucTut = "Low";
                else if (itut == 2)
                    sonucTut = "Middle";
                else
                    sonucTut = "High";
                if (sonucTut == test[i, 5])
                    dogru++;
                for (int j = 0; j < 5; j++)
                {
                    myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = test[i, j];
                    myRange.Select();
                }
                myRange = (Range)sheet1.Cells[StartRow + i, StartCol + 5];
                myRange.Value2 = sonucTut;
                myRange.Select();
                for (int j = 0; j < 4; j++)
                    tut[j] = 0;
            }

            return dogru;
        }
    }
}
