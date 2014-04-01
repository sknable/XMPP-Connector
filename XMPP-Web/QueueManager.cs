using System;
using System.Collections.Concurrent;
using System.Threading;

namespace XMPP_Web
{
    public class QueueManager : IDisposable
    {
        public volatile bool IsShutdown = false;
        public EventWaitHandle WaitHandle = new AutoResetEvent(false);
        private ConcurrentDictionary<String, ChatQueue> _queues = new ConcurrentDictionary<String, ChatQueue>();

        public void AddDirectQueue(String queue)
        {
            ChatQueue q = new ChatQueue(queue, "password");
            q.UseXMPPChatRoom = false;
            q.Connect();
            _queues.GetOrAdd(queue, q);
        }

        public void AddQueue(String queue)
        {
            ChatQueue q = new ChatQueue(queue, "password");
            q.UseXMPPChatRoom = true;
            q.Connect();
            _queues.GetOrAdd(queue, q);
        }

        public void CreateOutBoundSession(String queue, String jID, String message)
        {
            try
            {
                _queues[queue].CreateOutBoundSession(jID, message);
            }
            catch { }
        }

        public void DeleteQueue()
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void DoWork()
        {
            Boolean didFail = false;

            CScript.Instance.Initialize();
            CScript.Instance.UseSriptEngine = true;

            while (!IsShutdown)
            {
                WaitHandle.WaitOne(30000);

                if (!ChatSession.Ping())
                {
                    if (!didFail)
                    {
                        Logger.WriteLine("!!!Server Ping Failed - Shutting Down XMPP Queues!!!");
                        didFail = true;
                        foreach (var pair in _queues)
                        {
                            pair.Value.ShutDown();
                        }
                    }
                }
                else if (didFail)
                {
                    foreach (var pair in _queues)
                    {
                        pair.Value.Connect();
                    }
                    didFail = true;
                }
            }
        }

        public void SetInteropLogger()
        {
            Logger.StandAlone = true;
        }

        public void ShutDown()
        {
            IsShutdown = true;
            WaitHandle.Set();
        }

        public void UpdatePresence(String queue, String message, Boolean available)
        {
            try
            {
                _queues[queue].UpdatePresence(message, available);
            }
            catch { }
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                WaitHandle.Dispose();
            }
        }
    }
}