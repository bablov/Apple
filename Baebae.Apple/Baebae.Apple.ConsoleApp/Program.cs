using System;
using Baebaego.Practice.Genericity;
using static System.Console;
using static Baebae.Apple.ConsoleApp.SevenNewFeature;

namespace Baebae.Apple.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Hello World!");

            //var result = IsFirstSummerMonday(DateTime.Now);

            //WriteLine(result);

            #region ·ºÐÍ
            //var test = new Test<string, byte>("abc",255);

            //var s1 = new Sample<string>();
            //s1.Work("Hello");

            //var s2 = new Sample<DateTime>();
            //s2.Work(DateTime.Now);

            //var s3 = new Sample<decimal>();
            //s3.Work(0.33M);

            //var s4 = new Sample<float>();
            //s4.Work(11.954f);

            //var s5 = new Sample<byte>();
            //s5.Work(255);

            //var s6 = new Sample<uint>();
            //s6.Work(798652);

            Something1 v1 = new Something1();
            v1.Output(500, 99.88d);

            Something2<char, string> v3 = new Something2<char, string>();
            v3.Output('c', "cat");
            #endregion

            ReadKey();
        }
    }
}
