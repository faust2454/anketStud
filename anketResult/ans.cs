using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace anketStud
{
    class Ans
    {
        public int queid;
        public int ans;
        public int ansid;

        public Ans(int q, int a, int ai)
        {
            queid = q;
            ans = a;
            ansid = ai;
        }

        public Ans(string q, object a, string ai)
        {
            queid = Convert.ToInt32(q);
            ans = (int)a;
            ansid = Convert.ToInt32(ai);
        }
    }
}
