using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace anketEmp
{
    class Ans
    {
        public int queid;
        public int ans;
        public int ansid;
        public int[] ansarr;

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

        public Ans(object q, int[] a)
        {
            queid = Convert.ToInt32(q);
            ansarr = a;
        }
    }
}
