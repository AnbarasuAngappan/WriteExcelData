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
            string _ipaddress = Console.ReadLine();
            bool _response = modbusReader.OpenProtocol(_ipaddress, 502);
            var _reading = modbusReader.ReadHoldingregister("1", "3204", "5");
            Console.ForegroundColor = ConsoleColor.Blue;
            for (int i = 0; (i <= (5 - 1)); i++)
            {              
                Console.WriteLine(_reading[i] + "  ");
            }
            Console.ResetColor();
            Console.ReadLine();
        }
    }
}
