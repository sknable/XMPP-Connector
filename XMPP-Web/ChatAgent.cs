using agsXMPP;
using agsXMPP.protocol.client;
using agsXMPP.protocol.x.muc;
using System;
using System.Threading;

namespace XMPP_Web
{
    public class ChatAgent
    {
        #region Properties

        public Jid MyJID;
        private String _chatRoom;
        private Jid _customerJID;
        private String _mUCServerName = "muc.servername.com";
        private String _name;
        private String _password;
        private String _serverName = "servername";
        private XmppClientConnection _xmpp = null;
        private MucManager groupchat;

        public Boolean UseXMPPChatRoom { get; private set; }

        #endregion Properties

        #region Events

        public delegate void LoggedInEventHandler(object sender, EventArgs e);

        public delegate void MessageEventHandler(object sender, String message);

        public event LoggedInEventHandler LoggedInEvent;

        public event MessageEventHandler MessageEvent;

        private void RaiseLoggedInEvent()
        {
            LoggedInEventHandler tempEvent = LoggedInEvent;

            if (tempEvent != null)
            {
                tempEvent(this, new EventArgs());
            }
        }

        private void RaiseMessageEvent(String message)
        {
            MessageEventHandler tempEvent = MessageEvent;

            if (tempEvent != null)
            {
                tempEvent(this, message);
            }
        }

        #endregion Events

        public ChatAgent(String userName, String password, String chatRoom)
        {
            _name = userName;
            _password = password;
            _chatRoom = chatRoom;
            UseXMPPChatRoom = true;
            _xmpp = new XmppClientConnection(_serverName);
            _xmpp.Password = password;
            _xmpp.Username = userName;
            _xmpp.Resource = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            _xmpp.AutoResolveConnectServer = true;
            _xmpp.OnLogin += Xmpp_OnLogin;
            _xmpp.OnError += Xmpp_OnError;
            _xmpp.OnXmppConnectionStateChanged += Xmpp_OnXmppConnectionStateChanged;
            groupchat = new MucManager(_xmpp);

            _xmpp.RegisterAccount = true;

            MyJID = _xmpp.MyJID;
        }

        public ChatAgent(String userName, String password, Jid customerJID)
        {
            _name = userName;
            _password = password;
            _chatRoom = "";
            UseXMPPChatRoom = false;
            _xmpp = new XmppClientConnection(_serverName);
            _xmpp.Password = password;
            _xmpp.Username = userName;
            _xmpp.Resource = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            _xmpp.AutoResolveConnectServer = true;
            _customerJID = customerJID;
            _xmpp.OnLogin += Xmpp_OnLogin;
            _xmpp.OnError += Xmpp_OnError;
            _xmpp.OnXmppConnectionStateChanged += Xmpp_OnXmppConnectionStateChanged;
            _xmpp.OnPresence += Xmpp_OnPresence;
            _xmpp.OnMessage += Xmpp_OnMessage;
            _xmpp.RegisterAccount = true;
            groupchat = new MucManager(_xmpp);
        }

        public void Connect()
        {
            _xmpp.Open();
        }

        public void Disconnect()
        {
            _xmpp.Close();
        }

        public void SendMessage(String message)
        {
            if (UseXMPPChatRoom)
            {
                _xmpp.Send(new Message(new Jid(_chatRoom + "@" + _mUCServerName), MessageType.groupchat, message));
                CScript.Instance.MonitorMessage(message, _chatRoom + "@" + _mUCServerName, null, true);
            }
            else
            {
                _xmpp.Send(new Message(_customerJID, MessageType.chat, message));
                CScript.Instance.MonitorMessage(message, _customerJID.ToString(), null, true);
            }
        }

        private void Xmpp_OnError(object sender, Exception ex)
        {
            Logger.WriteLine(ex.ToString());
        }

        private void Xmpp_OnLogin(object sender)
        {
            Logger.WriteLine("Connected to XMPP");
            RaiseLoggedInEvent();

            if (UseXMPPChatRoom)
            {
                groupchat.JoinRoom(new Jid(_chatRoom + "@" + _mUCServerName), _name);
            }
        }

        private void Xmpp_OnMessage(object sender, Message msg)
        {
            //Offline Message
            if (msg.XDelay != null)
            {
            }
            //Typing indicator
            else if (msg.Body == null)
            {
            }
            else if (!UseXMPPChatRoom && msg.From.ToString() == _customerJID.ToString())
            {
                RaiseMessageEvent(msg.Body);
                CScript.Instance.MonitorMessage(msg.Body, msg.From.ToString(), null, true);
            }
        }

        private void Xmpp_OnPresence(object sender, Presence pres)
        {
            Logger.WriteLine("[" + _name + "]XMPP Presence: " + pres.Type.ToString() + " - " + pres.From.ToString());

            //Someone has added us to thier body list..
            if (pres.Type == PresenceType.subscribe)
            {
                _xmpp.PresenceManager.RefuseSubscriptionRequest(pres.From);
            }
        }

        private void Xmpp_OnXmppConnectionStateChanged(object sender, XmppConnectionState state)
        {
            Logger.WriteLine("XMPP State[" + _name + "]:" + state);
        }
    }
}