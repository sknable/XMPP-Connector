using System;
using System.Threading;
using XMPP_Web;

namespace XMPP_Connector
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var q = new QueueManager();
            Thread qThread = new Thread(q.DoWork);

            qThread.Start();
            q.AddQueue("cce-support");
            q.AddQueue("ctc-support");
            q.AddQueue("scp-support");
            q.AddDirectQueue("dip-support");
            //qThread.Join();

            Console.ReadLine();

            q.ShutDown();
        }
    }
}