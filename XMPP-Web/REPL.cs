using agsXMPP;
using agsXMPP.protocol.client;
using System;

namespace XMPP_Web
{
    public class REPL
    {
        public static Jid MyJid;
        public static ChatQueue MyQueue;
        public static XmppClientConnection Xmpp;

        static public void Echo(String message)
        {
            Xmpp.Send(new Message(MyJid, MessageType.chat, message));
        }
    }
}