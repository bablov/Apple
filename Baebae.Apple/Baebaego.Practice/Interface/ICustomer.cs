using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Baebaego.Practice.Interface
{
    public interface ICustomer
    {
        IEnumerable<IOrder> PreviousOrders { get; }

        DateTime DateJoined { get; }
        DateTime? LastOrder { get; }
        string Name { get; }
        IDictionary<DateTime,string> Reminders { get; }

        public decimal ComputeLoyaltyDiscount()
        {
            DateTime TwoYearsAgo = DateTime.Now.AddYears(-2);
            if((DateJoined < TwoYearsAgo) && (PreviousOrders.Count()>10))
            {
                return 0.10m;
            }
            return 0;
        }
    }

    public interface IOrder
    {
        DateTime Purchased { get; }
        decimal Cost { get; }
    }

    public class SampleCustomer : ICustomer
    {
        public SampleCustomer(string name, DateTime dataJoined) =>
            (name, DateJoined) = (name, dataJoined);

        private List<IOrder> allOrders = new List<IOrder>();

        public IEnumerable<IOrder> PreviousOrders => allOrders;

        public DateTime DateJoined {get;}

        public DateTime? LastOrder { get; private set; }

        public string Name { get; }

        public IDictionary<DateTime, string> reminders = new Dictionary<DateTime, string>();
        public IDictionary<DateTime, string> Reminders => reminders;

        public void AddOrder(IOrder order)
        {
            if (order.Purchased > (LastOrder ?? DateTime.MinValue))
                LastOrder = order.Purchased;
            allOrders.Add(order);
        }
    }
}
