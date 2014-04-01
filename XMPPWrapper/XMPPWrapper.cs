using System;
using System.Threading;
using XMPP_Web;

namespace XMPPWrapper
{
    public class XMPPWrapper
    {
        private QueueManager q;

        public void AddQueue(String queueName, Boolean direct)
        {
            if (direct)
            {
                q.AddDirectQueue(queueName);
            }
            else
            {
                q.AddQueue(queueName);
            }
        }

        public void CreateRequest(String queue, String jID, String message)
        {
            q.CreateOutBoundSession(queue, jID, message);
        }

        public void Shutdown()
        {
            q.ShutDown();
        }

        public void StartXMPP()
        {
            q = new QueueManager();
            q.SetInteropLogger();
            Thread qThread = new Thread(q.DoWork);

            qThread.Start();
            q.AddQueue("A-support");
            q.AddQueue("B-support");
            q.AddQueue("C-support");
            q.AddDirectQueue("D-support");
        }

        public void UpdatePresence(String queue, String message, Boolean available)
        {
            q.UpdatePresence(queue, message, available);
        }
    }
}
