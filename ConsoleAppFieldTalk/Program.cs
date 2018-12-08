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
                timeOut = int.Parse("");
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
            }
            else
            {
                //lblResult.Text = ("Could not open protocol, error was: " + BusProtocolErrors.getBusProtocolErrorText(res));
                Console.WriteLine("Could not open protocol, error was: " + BusProtocolErrors.getBusProtocolErrorText(res));
            }
            
        }
    }
}
