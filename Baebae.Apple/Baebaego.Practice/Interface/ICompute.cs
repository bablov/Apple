using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Baebaego.Practice.Interface
{
    interface ICompute
    {
        int Id { get; set; }
        string Name { get; set; }
        void Total();
        void Avg();
    }
}
