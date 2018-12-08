using System;
using System.Collections.Generic;using System.Linq;using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppTest
{
    class Program
    {
        ModBus modbusReader;
        static void Main(string[] args)
        {

            Program program = new Program();
            program.Test();
        }

        public void Test()
        {
            modbusReader = new ModBus();
            bool asss = modbusReader.OpenProtocal("192.168.5.178", 502);
        }
    }
}
