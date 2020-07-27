using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToCode.Enities
{
    public class Position
    {
        private int x;

        public int X
        {
            get { return x; }
            set { x = value; }
        }
        private int y;

        public int Y
        {
            get { return y; }
            set { y = value; }
        }

        public Position()
        {
            X = 0;
            Y = 0;
        }

        public Position(int x,int y)
        { 
           X = x - 1;
           Y = y - 1;
        }
    }
}
