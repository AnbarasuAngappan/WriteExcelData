using FieldTalk.Modbus.Master;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppFieldTalk
{
    class Program
    {
        private static MbusMasterFunctions myProtocol;
        private static int retryCnt;
        private static int pollDelay;
        private static int timeOut;
        private static int tcpPort;
        private static int res;

        static void Main(string[] args)
        {
            myProtocol = new MbusRtuOverTcpMasterProtocol();
            try
            {
                retryCnt = int.Parse("");
            }
            catch (Exception)
            {
                retryCnt = 0;
            }
            try
            {
                pollDelay = int.Parse("");
            }
            catch (Exception)
            {
                pollDelay = 0;
            }
            try
            {
                timeOut = int.Parse("1000");
            }
            catch (Exception)
            {
                timeOut = 1000;
            }
            try
            {
                tcpPort = int.Parse("502");
            }
            catch (Exception)
            {
                tcpPort = 502;
            }

            myProtocol.timeout = timeOut;
            myProtocol.retryCnt = retryCnt;
            myProtocol.pollDelay = pollDelay;

            ((MbusRtuOverTcpMasterProtocol)myProtocol).port = (short)tcpPort;
            res = ((MbusRtuOverTcpMasterProtocol)myProtocol).openProtocol("192.168.5.178");

            if ((res == BusProtocolErrors.FTALK_SUCCESS))
            {
                //lblResult.Text = ("Modbus/TCP port opened successfully with parameters: " + (txtHostName.Text + (", TCP port " + tcpPort)));
                Console.WriteLine("Modbus/TCP port opened successfully with parameters: " + ("192.168.5.178" + (", TCP port " + tcpPort)));
                Console.ReadLine();
                ReadHoldingRegisters();

            }
            else
            {
                //lblResult.Text = ("Could not open protocol, error was: " + BusProtocolErrors.getBusProtocolErrorText(res));
                Console.WriteLine("Could not open protocol, error was: " + BusProtocolErrors.getBusProtocolErrorText(res));
            }

        }

        static void ReadHoldingRegisters()
        {
            short[] writeVals = new short[125];
            short[] readVals = new short[125];
            int slave;
            int startWrReg;
            int numWrRegs;
            int startRdReg;
            int numRdRegs;
            int i;
            int res;
            int startCoil;
            int numCoils;
            bool[] coilVals = new bool[2000];
            try
            {
                try
                {
                    slave = int.Parse("1");
                }
                catch (Exception)
                {
                    slave = 1;
                }
                try
                {
                    startCoil = int.Parse("");
                }
                catch (Exception)
                {
                    startCoil = 1;
                }
                try
                {
                    numCoils = int.Parse("5");
                }
                catch (Exception)
                {
                    numCoils = 1;
                }
                try
                {
                    startRdReg = int.Parse("3204");
                }
                catch (Exception)
                {
                    startRdReg = 1;
                }
                try
                {
                    startWrReg = int.Parse("");
                }
                catch (Exception)
                {
                    startWrReg = 1;
                }
                try
                {
                    numWrRegs = int.Parse("");
                }
                catch (Exception)
                {
                    numWrRegs = 1;
                }
                try
                {
                    numRdRegs = int.Parse("5");
                }
                catch (Exception)
                {
                    numRdRegs = 1;
                }
                try
                {
                    writeVals[0] = Int16.Parse(null);
                    writeVals[1] = Int16.Parse(null);
                    writeVals[2] = Int16.Parse(null);
                    writeVals[3] = Int16.Parse(null);
                    writeVals[4] = Int16.Parse(null);
                    writeVals[5] = Int16.Parse(null);
                    writeVals[6] = Int16.Parse(null);
                    writeVals[7] = Int16.Parse(null);
                    coilVals[0] = (writeVals[0] != 0);
                    coilVals[1] = (writeVals[1] != 0);
                    coilVals[2] = (writeVals[2] != 0);
                    coilVals[3] = (writeVals[3] != 0);
                    coilVals[4] = (writeVals[4] != 0);
                    coilVals[5] = (writeVals[5] != 0);
                    coilVals[6] = (writeVals[6] != 0);
                    coilVals[7] = (writeVals[7] != 0);
                }
                catch (Exception)
                {
                }


                res = myProtocol.readMultipleRegisters(slave, startRdReg, readVals, numRdRegs);
                //lblResult2.Text = ("Result: " + (BusProtocolErrors.getBusProtocolErrorText(res) + "\r\n"));
                string a = ("Result: " + (BusProtocolErrors.getBusProtocolErrorText(res) + "\r\n"));
                if ((res == BusProtocolErrors.FTALK_SUCCESS))
                {
                    //lblReadValues.Text = "";
                    for (i = 0; (i <= (numRdRegs - 1)); i++)
                    {
                        //lblReadValues.Text = (a + (readVals[i] + "  "));
                        Console.WriteLine(a + readVals[i] + "  ");
                    }
                    Console.ReadLine();
                }
            }
            catch
            {

            }

        }
    }
}
