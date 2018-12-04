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

            using (TcpClient client = new TcpClient("192.168.5.178", 502))
            {
                ModbusIpMaster master = ModbusIpMaster.CreateIp(client);

                // read five input values
                ushort startAddress = 3204;//100;
                ushort numInputs = 5;
               // bool[] inputs = master.ReadInputs(startAddress, numInputs);
                ushort[] PO = master.ReadHoldingRegisters(startAddress, numInputs);

                for (int i = 0; i < numInputs; i++)
                {
                    Console.WriteLine($"Input {(startAddress + i)}={""}");
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
