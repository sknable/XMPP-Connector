using ChatScript;
using System;

namespace XMPP_Web
{
    internal class CScript
    {
        public Boolean UseSriptEngine;
        private static CScript _instance;
        private CSScriptEngine _chatScriptEngine = new CSScriptEngine();

        private CScript()
        {
        }

        public static CScript Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new CScript();
                }
                return _instance;
            }
        }

        public void ChatInQueueUpdate(int count, ChatQueue queue)
        {
            if (UseSriptEngine)
            {
                try
                {
                    _chatScriptEngine.Script.OnChatInQueueUpdate(count, queue);
                }
                catch { }
            }
        }

        public void ChatSessionEnded(String customer, ChatQueue queue, String agentName)
        {
            if (UseSriptEngine)
            {
                try
                {
                    _chatScriptEngine.Script.OnChatSessionEnded(customer, queue, agentName);
                }
                catch { }
            }
        }

        public void ChatSessionStarted(String customer, ChatQueue queue, String agentName)
        {
            if (UseSriptEngine)
            {
                try
                {
                    _chatScriptEngine.Script.OnChatSessionStarted(customer, queue, agentName);
                }
                catch { }
            }
        }

        public void Initialize()
        {
            if (!_chatScriptEngine.LoadScriptFile(@"Chat.cs"))
            {
                //For Demo
                if (!_chatScriptEngine.LoadScriptFile(@"c:\demo\site\XMPP\Chat.cs"))
                {
                    Logger.WriteLine("Cant load Chat.cs");
                }
            }

            try
            {
                _chatScriptEngine.Script.Initialize();
            }
            catch { }
        }

        public void MonitorMessage(String message, String from, ChatQueue queue, Boolean isAgent)
        {
            if (UseSriptEngine)
            {
                try
                {
                    _chatScriptEngine.Script.OnMonitorMessage(message, from, queue, isAgent);
                }
                catch { }
            }
        }

        public Boolean NewDirectMessage(String message, String from, ChatQueue queue)
        {
            if (UseSriptEngine)
            {
                try
                {
                    return _chatScriptEngine.Script.OnNewDirectMessage(message, from, queue);
                }
                catch { return true; }
            }
            else
            {
                return true;
            }
        }

        public Boolean OutBoundMessageConfirmation(String message, String from, ChatQueue queue)
        {
            if (UseSriptEngine)
            {
                try
                {
                    return _chatScriptEngine.Script.OnOutBoundMessageConfirmation(message, from, queue);
                }
                catch { return true; }
            }
            else
            {
                return true;
            }
        }

        public void Run(String line)
        {
            _chatScriptEngine.Run(line);
        }
    }
}
