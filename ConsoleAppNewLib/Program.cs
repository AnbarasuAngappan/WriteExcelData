using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using Modbus.Data;
using Modbus.Device;

namespace ConsoleAppNewLib
{
    class Program
    {
        static void Main(string[] args)
        {
          
                using (TcpClient client = new TcpClient("127.0.0.1", 502))
                {
                    ModbusIpMaster master = ModbusIpMaster.CreateIp(client);

                    // read five input values
                    ushort startAddress = 100;
                    ushort numInputs = 5;
                    bool[] inputs = master.ReadInputs(startAddress, numInputs);

                    for (int i = 0; i < numInputs; i++)
                    {
                        Console.WriteLine($"Input {(startAddress + i)}={(inputs[i] ? 1 : 0)}");
                    }
                }

                // output: 
                // Input 100=0
                // Input 101=0
                // Input 102=0
                // Input 103=0
                // Input 104=0
           
        }
    }
}
