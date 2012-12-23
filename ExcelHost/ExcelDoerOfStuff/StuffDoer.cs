using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDoerOfStuff
{
    public class StuffDoer
    {
        public Guid Id { get; private set; }
        private static StuffDoer _instance;

        private StuffDoer()
        {
            Id = Guid.NewGuid();
        }

        public static StuffDoer GetInstance()
        {
            return _instance ?? (_instance = new StuffDoer());
        }
    }
}