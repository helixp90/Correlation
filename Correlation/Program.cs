using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Correlation
{
    class Program
    {
        static void Main(string[] args)
        {
            
            List<string> pm1 = new List<string>(), pm2 = new List<string>(), pm3 = new List<string>(), pm4 = new List<string>(), so1 = new List<string>(), so2 = new List<string>(), so3 = new List<string>(), so4 = new List<string>(), no1 = new List<string>(), no2 = new List<string>(), no3 = new List<string>(), no4 = new List<string>(), o1 = new List<string>(), o2 = new List<string>(), o3 = new List<string>(), o4 = new List<string>(), target = new List<string>();

            using (StreamReader sr = new StreamReader("ProcessedData.csv"))
            {


                while (!sr.EndOfStream) // read Excel File till end value
                {
                    var line = sr.ReadLine();
                    
                    var values = line.Split(','); // split columns via commas

                    pm1.Add(values[2]);
                    so1.Add(values[3]);
                    no1.Add(values[4]);
                    o1.Add(values[5]);

                    pm2.Add(values[6]);
                    so2.Add(values[7]);
                    no2.Add(values[8]);
                    o2.Add(values[9]);

                    pm3.Add(values[10]);
                    so3.Add(values[11]);
                    no3.Add(values[12]);
                    o3.Add(values[13]);

                    pm4.Add(values[14]);
                    so4.Add(values[15]);
                    no4.Add(values[16]);
                    o4.Add(values[17]);

                    target.Add(values[18]);
                }

            }

            pm1.RemoveAt(0);
            target.RemoveAt(0);

            List<double> myList = pm1.ConvertAll(item => double.Parse(item));
            List<double> myList2 = target.ConvertAll(item => double.Parse(item));

            var application = new Application();

            var worksheetFunction = application.WorksheetFunction;

            var result = worksheetFunction.Correl(myList.ToArray(), myList2.ToArray());

            Console.Write(result);

            //pm1.ForEach(Console.WriteLine);

            //Console.WriteLine(pm1 + "\n" + so1);
            Console.ReadKey();

           }
        }
    }

