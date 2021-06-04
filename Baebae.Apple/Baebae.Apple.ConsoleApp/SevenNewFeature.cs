using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Baebae.Apple.ConsoleApp
{
    public static class SevenNewFeature
    {
        public static void PatternMatching()
        {
            int sum = 0;
            object input = 1;
            if(input is int count)
            {
                sum += count;
            }

            Console.WriteLine(sum);
        }

        public static bool IsFirstSummerMonday(DateTime date) => date is { Month: 6, Day: <= 7, DayOfWeek: DayOfWeek.Monday };
    }
}
