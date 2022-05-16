using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("Введите название файла");
                string pathXlsxFile = Console.ReadLine();

                pathXlsxFile = Directory.GetCurrentDirectory() + "/" + pathXlsxFile + ".xlsx";

                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(pathXlsxFile, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                int numCol = 1;
                Range usedColumn = ObjWorkSheet.UsedRange.Columns[numCol];
                System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
                string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                List<double> fullData = new List<double>();
                for (int i = 0; i < strArray.Length; i++)
                {
                    fullData.Add(double.Parse(strArray[i]));
                }

                ObjExcel.Quit();


                FirstTenItems(fullData, 10, 10);
                FirstTenItems(fullData, 10, 30);
                FirstTenItems(fullData, 10, 50);

                FirstTenItems(fullData, 20, 10);
                FirstTenItems(fullData, 20, 30);
                FirstTenItems(fullData, 20, 50);


                FullMDData(fullData);
            }
        }


        static void FullMDData(List<double> data)
        {
            
            //мат ожидание
            double MatOjid = 0;
            for (int i = 0; i < data.Count; i++)
                MatOjid += data[i];
            MatOjid /= 44;
            Console.WriteLine("Матиматическое ожидание обычным методом всей выборки: " + MatOjid.ToString());

            //дисперсия
            double disp = 0.0;
            for (int i = 0; i < data.Count; i++)
            {
                disp += (data[i] - MatOjid) * (data[i] - MatOjid);
            }
            disp /= 43;
            disp = Math.Round(disp, 2, MidpointRounding.AwayFromZero);
            Console.WriteLine("Дисперсия обычным методом всей выборки: " + disp.ToString());
        }

        static void FirstTenItems(List<double> data, int size, int povtor)
        {
            Console.WriteLine("Число повторений: " + povtor.ToString() + ". Размер данных: " + size.ToString());
            //мат ожидание
            double MatOjid = 0;
            for (int i = 0; i < size; i++)
                MatOjid += data[i];
            MatOjid /= size;
            Console.WriteLine("Матиматическое ожидание обычным методом: " + MatOjid.ToString());

            //дисперсия
            double disp = 0.0;
            for (int i = 0; i < size; i++)
            {
                disp += (data[i] - MatOjid) * (data[i] - MatOjid);
            }
            disp /= size - 1;
            disp = Math.Round(disp, 2, MidpointRounding.AwayFromZero);
            Console.WriteLine("Дисперсия обычным методом: " + disp.ToString());

            var rand = new Random();

            //bootstrap мат ожидание 10
            List<double> MsoZvezdoiOtI = new List<double>();
            List<List<int>> list_randomov = new List<List<int>>();
            for (int i = 0; i < povtor;)
            {
                List<int> randomList = new List<int>();
                for (int j = 0; j < size; j++)
                {
                    randomList.Add(rand.Next(4));

                }
                if (randomList.Sum() != size)
                {
                    continue;
                }

                list_randomov.Add(randomList);

                double sumProizv = 0;
                for (int j = 0; j < size; j++)
                {
                    sumProizv += (randomList[j] * data[j]);
                }

                sumProizv /= size;
                MsoZvezdoiOtI.Add(sumProizv);
                i++;
            }
            double MsoZvezdoi = 0;
            for (int i = 0; i < povtor; i++)
            {
                MsoZvezdoi += MsoZvezdoiOtI[i];
            }
            MsoZvezdoi /= povtor;
            double deltaM = MsoZvezdoi - MatOjid;
            Console.WriteLine("M*: " + MsoZvezdoi.ToString());
            Console.WriteLine("deltaM: " + deltaM.ToString());


            //D*


            List<double> DsoZvezdoiOtI = new List<double>();
            for (int i = 0; i < povtor; i++)
            {
                double sumDOtI = 0;
                for (int j = 0; j < size; j++)
                {
                    sumDOtI += Math.Pow(data[j] - MsoZvezdoiOtI[i], 2) * list_randomov[i][j];
                }

                DsoZvezdoiOtI.Add(sumDOtI/(size - 1));
                
            }
            double sumDotILista = 0;
            for (int j = 0; j < povtor; j++)
                sumDotILista += DsoZvezdoiOtI[j];

            double DsoZvezdoi = sumDotILista / povtor;

            DsoZvezdoi = Math.Round(DsoZvezdoi, 2, MidpointRounding.AwayFromZero);
            double deltaD = DsoZvezdoi - disp;

            Console.WriteLine("D*: " + DsoZvezdoi.ToString());
            Console.WriteLine("deltaD: " + deltaD.ToString());


            //DM* DD*

            double sumDM = 0, sumDD = 0;
            for(int i = 0; i < povtor; i++)
            {
                sumDM += Math.Pow(MsoZvezdoiOtI[i] - MsoZvezdoi,2);
                sumDD += Math.Pow(DsoZvezdoiOtI[i] - DsoZvezdoi,2);
            }
            double DM = sumDM / (povtor - 1), DD = sumDD / (povtor - 1);
            Console.WriteLine("DM*: " + DM.ToString());
            Console.WriteLine("DD*: " + DD.ToString());
            Console.WriteLine("\n");

        }
    }
}
