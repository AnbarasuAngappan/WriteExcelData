using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using NModbus.Device;

namespace NModbus_Dmitry
{
    class Program
    {
        static void Main(string[] args)
        {
            using (TcpClient client = new TcpClient("192.168.5.178", 502))
            {
                
                //ModbusIpMaster  

                //ModbusIpMaster master = ModbusIpMaster.CreateIp(client);
               
                // read five input values
                //ushort startAddress = 3204;
                //ushort numInputs = 2;
                ////bool[] inputs = master.ReadInputs(startAddress, numInputs);
                //ushort[] outHolding = master.ReadHoldingRegisters(startAddress, numInputs);

                //for (int i = 0; i < numInputs; i++)
                //{
                //    Console.WriteLine($"Input {(startAddress + i)}={"NULL"}");
                //}
            }
        }
    }
}
