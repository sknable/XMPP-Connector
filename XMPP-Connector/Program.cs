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
            q.AddQueue("A-support");
            q.AddQueue("B-support");
            q.AddQueue("C-support");
            q.AddDirectQueue("D-support");


            Console.ReadLine();

            q.ShutDown();
        }
    }
}
