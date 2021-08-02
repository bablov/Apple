using System;
using System.Linq;

namespace Linq
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var test1 = Enumerable.Range(1, 10).ElementAt(^2);  // returns 9
            var test2 = Enumerable.Range(1, 10).Take(^2..);  // returns [9,10]
            var test3 = Enumerable.Range(1, 10).Take(..2);  // returns [1,2]
            var test4 = Enumerable.Range(1, 10).Take(2..4);  // returns [3,4]

            var test5 = Enumerable.Range(1, 10).DistinctBy(x => x % 3);

            var first = new (string Name, int Age)[] { ("Francis", 20), ("Lindsey", 30), ("Ashley", 40) };
            var second = new (string Name, int Age)[] { ("Claire", 30), ("Pat", 30), ("Drew", 33) };
            var test6 = first.UnionBy(second,person=>person.Age).Select(x => $"{x.Name},{x.Age}");
            var test7 = second.UnionBy(first, person=>person.Age).Select(x => $"{x.Name},{x.Age}");

            Console.ReadKey();
        }
    }
}
