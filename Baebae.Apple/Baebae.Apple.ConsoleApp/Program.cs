using System;
using static System.Console;
using static Baebae.Apple.ConsoleApp.SevenNewFeature;

namespace Baebae.Apple.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteLine("Hello World!");

            var result = IsFirstSummerMonday(DateTime.Now);

            WriteLine(result);

            ReadKey();
        }
    }
}
