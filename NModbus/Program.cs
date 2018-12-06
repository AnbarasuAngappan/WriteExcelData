using Modbus.Device;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace NModbus
{
    class Program
    {
        static void Main(string[] args)
        {
            using (TcpClient client = new TcpClient("192.168.5.178", 502))
            {
                ModbusIpMaster master = ModbusIpMaster.CreateIp(client);

                // read five input values
                ushort startAddress = 3200;
                ushort numInputs = 25;
                //bool[] inputs = master.ReadInputs(startAddress, numInputs);
                ushort[] outHolding = master.ReadHoldingRegisters(startAddress, numInputs);

                for (int i = 0; i < numInputs; i++)
                {
                    Console.WriteLine($"Input {(startAddress + i)}={"NULL"}");
                }
            }
        }
    }
}
