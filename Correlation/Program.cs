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

            List<double> fpm1 = new List<double>(), fpm2 = new List<double>(), fpm3 = new List<double>(), fpm4 = new List<double>(), fso1 = new List<double>(), fso2 = new List<double>(), fso3 = new List<double>(), fso4 = new List<double>(), fno1 = new List<double>(), fno2 = new List<double>(), fno3 = new List<double>(), fno4 = new List<double>(), fo1 = new List<double>(), fo2 = new List<double>(), fo3 = new List<double>(), fo4 = new List<double>(), ftarget = new List<double>();


            using (StreamReader sr = new StreamReader("ProcessedData.csv"))
            {


                while (!sr.EndOfStream) // read Excel File till end value
                {
                    var line = sr.ReadLine();
                    
                    var values = line.Split(','); // split columns via commas

                    pm1.Add(values[2]);         // adds columns to lists, including headers
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

            simplify(pm1,so1, no1, o1, pm2, so2, no2, o2, pm3, so3, no3, o3, pm4, so4, no4, o4, target); // simplify Lists for correlation

            
            fpm1 = pm1.ConvertAll(item => double.Parse(item));
            fso1 = so1.ConvertAll(item => double.Parse(item));
            fno1 = no1.ConvertAll(item => double.Parse(item));
            fo1 = o1.ConvertAll(item => double.Parse(item));

            fpm2 = pm2.ConvertAll(item => double.Parse(item));
            fso2 = so2.ConvertAll(item => double.Parse(item));
            fno2 = no2.ConvertAll(item => double.Parse(item));
            fo2 = o2.ConvertAll(item => double.Parse(item));

            fpm3 = pm3.ConvertAll(item => double.Parse(item));
            fso3 = so3.ConvertAll(item => double.Parse(item));
            fno3 = no3.ConvertAll(item => double.Parse(item));
            fo3 = o3.ConvertAll(item => double.Parse(item));

            fpm4 = pm4.ConvertAll(item => double.Parse(item));
            fso4 = so4.ConvertAll(item => double.Parse(item));
            fno4 = no4.ConvertAll(item => double.Parse(item));
            fo4 = o4.ConvertAll(item => double.Parse(item));

            ftarget = target.ConvertAll(item => double.Parse(item));

            var application = new Application();

            var worksheetFunction = application.WorksheetFunction;

            Console.WriteLine("1st 4 columns: \n\nPM1: " + worksheetFunction.Correl(fpm1.ToArray(), ftarget.ToArray()) + "\nSO1: " + worksheetFunction.Correl(fso1.ToArray(), ftarget.ToArray()) + "\nNO1: " + worksheetFunction.Correl(fno1.ToArray(), ftarget.ToArray()) + "\nO1: " + worksheetFunction.Correl(fo1.ToArray(), ftarget.ToArray()) + "\n\n");

            Console.WriteLine("2nd 4 columns: \nPM2: " + worksheetFunction.Correl(fpm2.ToArray(), ftarget.ToArray()) + "\nSO2: " + worksheetFunction.Correl(fso2.ToArray(), ftarget.ToArray()) + "\nNO2: " + worksheetFunction.Correl(fno2.ToArray(), ftarget.ToArray()) + "\nO2: " + worksheetFunction.Correl(fo2.ToArray(), ftarget.ToArray()) + "\n\n");

            Console.WriteLine("3rd 4 columns: \nPM1: " + worksheetFunction.Correl(fpm3.ToArray(), ftarget.ToArray()) + "\nSO3: " + worksheetFunction.Correl(fso3.ToArray(), ftarget.ToArray()) + "\nNO3: " + worksheetFunction.Correl(fno3.ToArray(), ftarget.ToArray()) + "\nO3: " + worksheetFunction.Correl(fo3.ToArray(), ftarget.ToArray()) + "\n\n");

            Console.WriteLine("Last 4 columns: \nPM4: " + worksheetFunction.Correl(fpm4.ToArray(), ftarget.ToArray()) + "\nSO4: " + worksheetFunction.Correl(fso4.ToArray(), ftarget.ToArray()) + "\nNO4: " + worksheetFunction.Correl(fno4.ToArray(), ftarget.ToArray()) + "\nO4: " + worksheetFunction.Correl(fo4.ToArray(), ftarget.ToArray()) + "\n\n");

            //var result = worksheetFunction.Correl(myList.ToArray(), myList2.ToArray());


            //Console.WriteLine(pm1 + "\n" + so1);
            Console.ReadKey();

           }

        static void simplify(List<string> c, List<string> d, List<string> e, List<string> f, List<string> g, List<string> h, List<string> i, List<string> j, List<string> k, List<string> l, List<string> m, List<string> n, List<string> o, List<string> p, List<string> q, List<string> r, List<string> s)
        {
            c.RemoveAt(0);  // remove headers
            d.RemoveAt(0);
            e.RemoveAt(0);
            f.RemoveAt(0);
            g.RemoveAt(0);
            h.RemoveAt(0);
            i.RemoveAt(0);
            j.RemoveAt(0);
            k.RemoveAt(0);
            l.RemoveAt(0);
            m.RemoveAt(0);
            n.RemoveAt(0);
            o.RemoveAt(0);
            p.RemoveAt(0);
            q.RemoveAt(0);
            r.RemoveAt(0);
            s.RemoveAt(0);
        }

    }

 }

