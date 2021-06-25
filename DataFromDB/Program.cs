using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataFromDB
{
    class Program
    {
        static void Main(string[] args)
        {            
            IrbisHandler irbisHandler = new IrbisHandler();            
            irbisHandler.Perform();
        }

       
    }
}
