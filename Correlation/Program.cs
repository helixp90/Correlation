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

