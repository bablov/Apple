using System;

namespace Baebae.Apple.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var result = SevenNewFeature.IsFirstSummerMonday(DateTime.Now);

            Console.WriteLine(result);

            Console.ReadKey();
        }
    }
}
