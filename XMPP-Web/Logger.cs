using System;

namespace XMPP_Web
{
    internal class Logger
    {
        public static Boolean StandAlone = true;
        //private static ConnectionType interopCom = null;

        public static void WriteLine(String log)
        {
            Console.WriteLine(log);
        }
    }
}