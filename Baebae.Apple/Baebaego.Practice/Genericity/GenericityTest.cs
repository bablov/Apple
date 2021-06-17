using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Console;

namespace Baebaego.Practice.Genericity
{
    public class Test<T,F>
    {
        public Test(T a,F b)
        {
            WriteLine(a.GetType().Name);
            WriteLine(b.GetType().Name);
        }
    }

    public class Sample<K>
    {
        public void Work(K p)
        {
            WriteLine($"{p.GetType().FullName,-20}:{p}");
        }
    }

    public interface ITest<P,Q>
    {
        void Output(P x, Q y);
    }

    public class Something1:ITest<int,double>
    {
        public void Output(int x,double y)
        {
            WriteLine($"{x.GetType()}-{x}");
            WriteLine($"{y.GetType()}-{y}");
        }
    }
    public class Something2<J,K> : ITest<J,K>
    {
        public void Output(J x, K y)
        {
            WriteLine($"{x.GetType()}-{x}");
            WriteLine($"{y.GetType()}-{y}");
        }
    }

}
