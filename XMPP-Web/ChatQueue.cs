using agsXMPP;
using agsXMPP.protocol.client;
using agsXMPP.protocol.x.muc;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace XMPP_Web
{
    public class ChatQueue
    {
        #region CallBack Classes

        internal class CustomerChat
        {
            public Jid ChatRoom;
            public Jid Customer;
        }

        internal class GroupChatRoom
        {
            public Jid Chatjid;
            public Jid Customerjig;
        }

        internal class OutBoundChat
        {
            public Jid Customer;
        }

        #endregion CallBack Classes

        #region Proprties

        //ToDo: Make these a interlock
        public volatile int AvgWaitTime;

        public volatile int ChatInProgress;

        public volatile int ChatInQueue;

        private static String _mUCServerName = "muc.steveknable.com";

        private static String _serverName = "steveknable.com";

        private MucManager _groupChat;
        private String _password;

        private XmppClientConnection _xmpp = null;

        //Current XMPP Chat Rooms
        private ConcurrentDictionary<String, CustomerChat> _xmppChatRooms = new ConcurrentDictionary<String, CustomerChat>();

        // OutBound XMPP Sessions
        private ConcurrentDictionary<String, OutBoundChat> _xmppOutBoundSessions = new ConcurrentDictionary<String, OutBoundChat>();

        //Current Web Chat HTTP Sessions
        private ConcurrentDictionary<String, ChatSession> _xmppSessions = new ConcurrentDictionary<String, ChatSession>();

        public String _name { get; private set; }

        public Boolean UseXMPPChatRoom { get; set; }

        #endregion Proprties

        #region public methods

        public ChatQueue(String queueName, String password)
        {
            _name = queueName;
            _password = password;
            _xmpp = new XmppClientConnection(_serverName);
            _xmpp.Password = password;
            _xmpp.Username = queueName;
            //Needed for SVN record
            _xmpp.AutoResolveConnectServer = true;
            _xmpp.Resource = "CCEXMPP";
            //All XMPP events
            _xmpp.OnLogin += Xmpp_OnLogin;
            _xmpp.OnMessage += Xmpp_OnMessage;
            _xmpp.OnError += Xmpp_OnError;
            _xmpp.OnXmppConnectionStateChanged += Xmpp_OnXmppConnectionStateChanged;
            _xmpp.OnPresence += Xmpp_OnPresence;

            //Always try to create account...probably bad
            _xmpp.RegisterAccount = true;
        }

        public void Connect()
        {
            ChatInProgress = 0;
            ChatInQueue = 0;
            AvgWaitTime = 0;
            _xmpp.Open();
        }

        public void CreateOutBoundSession(String jID, String message)
        {
            OutBoundChat chat = new OutBoundChat();
            chat.Customer = new Jid(jID);
            _xmppOutBoundSessions.GetOrAdd(jID, chat);

            _xmpp.Send(new Message(chat.Customer, MessageType.chat, message));
        }

        public void SendChatMessageAsQueue(String message, String from)
        {
            if (from.Contains(_mUCServerName))
            {
                _xmpp.Send(new Message(new Jid(from), MessageType.groupchat, message));
            }
            else
            {
                _xmpp.Send(new Message(new Jid(from), MessageType.chat, message));
            }
        }

        public void SendChatMessageAsQueue(String plainTxtMessage, String htmlMessage, String from)
        {
            Message message;
            if (from.Contains(_mUCServerName))
            {
                message = new Message(new Jid(from), MessageType.groupchat, plainTxtMessage);
            }
            else
            {
                message = new Message(new Jid(from), MessageType.chat, plainTxtMessage);
            }

            message.Html = new agsXMPP.protocol.extensions.html.Html();
            message.Html.InnerXml = htmlMessage;

            _xmpp.Send(message);
        }

        public void UpdatePresence(String message, Boolean available)
        {
            Presence p;

            if (available)
            {
                p = new Presence(ShowType.chat, message);
            }
            else
            {
                p = new Presence(ShowType.dnd, message);
            }
            p.Priority = 5;
            _xmpp.Send(p);
        }

        #endregion public methods

        #region events

        public delegate void ChatSessionCanceledHandler(object sender, EventArgs e);

        public delegate void ChatSessionEndedHandler(object sender, EventArgs e);

        public delegate void ChatSessionStartedHandler(object sender, EventArgs e);

        public delegate void ChatSessionTakenHandler(object sender, EventArgs e);

        public event ChatSessionCanceledHandler ChatSessionCancelEvent;

        public event ChatSessionEndedHandler ChatSessionEndedEvent;

        public event ChatSessionStartedHandler ChatSessionStartedEvent;

        public event ChatSessionTakenHandler ChatSessionTakenEvent;

        private void RaiseChatSessionCancelEvent()
        {
            ChatSessionCanceledHandler tempEvent = ChatSessionCancelEvent;
            ChatInQueue--;
            CScript.Instance.ChatInQueueUpdate(ChatInQueue, this);

            if (tempEvent != null)
            {
                tempEvent(this, new EventArgs());
            }
        }

        private void RaiseChatSessionEnded()
        {
            ChatSessionEndedHandler tempEvent = ChatSessionEndedEvent;
            ChatInProgress--;

            if (tempEvent != null)
            {
                tempEvent(this, new EventArgs());
            }
        }

        private void RaiseChatSessionStartedEvent()
        {
            ChatInQueue++;
            ChatSessionStartedHandler tempEvent = ChatSessionStartedEvent;
            CScript.Instance.ChatInQueueUpdate(ChatInQueue, this);

            if (tempEvent != null)
            {
                tempEvent(this, new EventArgs());
            }
        }

        private void RaiseChatSessionTakenEvent()
        {
            ChatSessionTakenHandler tempEvent = ChatSessionTakenEvent;
            ChatInQueue--;
            ChatInProgress++;
            CScript.Instance.ChatInQueueUpdate(ChatInQueue, this);

            if (tempEvent != null)
            {
                tempEvent(this, new EventArgs());
            }
        }

        #endregion events

        #region XMPP Event Handlers

        private void Xmpp_OnCreateRoom(object sender, IQ iq, object data)
        {
            Thread.Sleep(100);
            GroupChatRoom room = (GroupChatRoom)data;
            _groupChat.JoinRoom(room.Chatjid, _name);
            Thread.Sleep(100);
            _groupChat.Invite(room.Customerjig, room.Chatjid);
        }

        private void Xmpp_OnDestroyRoom(object sender, IQ iq, object data)
        {
        }

        private void Xmpp_OnError(object sender, Exception ex)
        {
            Logger.WriteLine(ex.ToString());
        }

        private void Xmpp_OnLogin(object sender)
        {
            Logger.WriteLine("[" + _name + "]Connected to XMPP");
            _groupChat = new MucManager(_xmpp);
        }

        private void Xmpp_OnMessage(object sender, Message msg)
        {
            //Offline Message
            if (msg.XDelay != null)
            {
                Logger.WriteLine("[Offline Msg" + _name + "]" + msg.From.Resource + ": " + msg.Body + " To: " + msg.To);
            }
            //Typing indicator
            else if (msg.Body != null)
            {
                //C# REPL -- Sample
                if (msg.Body.StartsWith("!"))
                {
                    REPL.Xmpp = _xmpp;
                    REPL.MyJid = msg.From;
                    REPL.MyQueue = this;
                    CScript.Instance.Run(msg.Body.Remove(0, 2));
                }
                else if (msg.From.ToString().Contains(_mUCServerName))
                {
                    HandleMessageInChatRoom(msg);
                }
                else
                {
                    HandleDirectMessage(msg);
                }
            }
            else
            {
                Logger.WriteLine("[" + _name + "]Chat State: " + msg.Chatstate.ToString());
            }
        }

        private void Xmpp_OnPresence(object sender, Presence pres)
        {
            Logger.WriteLine("[" + _name + "]XMPP Presence: " + pres.Type.ToString() + " - " + pres.From.ToString());

            //Someone has added us to thier body list..
            if (pres.Type == PresenceType.subscribe)
            {
                _xmpp.PresenceManager.ApproveSubscriptionRequest(pres.From);
                //subscribe to get presence
                _xmpp.PresenceManager.Subscribe(pres.From);
            }
            //Someone Removed us...
            else if (pres.Type == PresenceType.unsubscribe)
            {
                _xmpp.PresenceManager.Unsubscribe(pres.From);
            }
            else if (pres.Type == PresenceType.unavailable)
            {
                //Someone left a chat that we are in
                if (UseXMPPChatRoom && pres.From.ToString().Contains(_mUCServerName))
                {
                    try
                    {
                        ChatSession session = _xmppSessions[pres.From.User];
                        if (session.SessionState == ChatState.CONNECTED)
                        {
                            RaiseChatSessionCancelEvent();
                        }
                        session.DisconnectChat();
                        Session_DisconnectedEvent(session);
                    }
                    catch { }
                }
            }
        }

        private void Xmpp_OnXmppConnectionStateChanged(object sender, XmppConnectionState state)
        {
            Logger.WriteLine("[" + _name + "]XMPP State: " + state);
        }

        #endregion XMPP Event Handlers

        private void HandleDirectMessage(Message msg)
        {
            Logger.WriteLine("[Direct-" + _name + "]" + msg.From + ": " + msg.Body);

            if (!UseXMPPChatRoom)
            {
                if (_xmppOutBoundSessions.ContainsKey(msg.From.Bare) && !_xmppSessions.ContainsKey(msg.From.ToString()))
                {
                    if (CScript.Instance.OutBoundMessageConfirmation(msg.Body, msg.From.ToString(), this))
                    {
                        CreateChatSession(msg.From.ToString(), true);
                    }
                }
                //New Conversation
                else if (!_xmppSessions.ContainsKey(msg.From.ToString()))
                {
                    if (CScript.Instance.NewDirectMessage(msg.Body, msg.From.ToString(), this))
                    {
                        CreateChatSession(msg.From.ToString(), false);
                    }
                }
            }
            else
            {
                if (_xmppOutBoundSessions.ContainsKey(msg.From.Bare) && !_xmppChatRooms.ContainsKey(msg.From.ToString()))
                {
                    if (CScript.Instance.OutBoundMessageConfirmation(msg.Body, msg.From.ToString(), this))
                    {
                        String chatRoomName = CreateChatRoom(msg.From);
                        CreateChatSession(chatRoomName, msg.From.ToString(), true);
                    }
                }
                //New Conversation
                else if (!_xmppChatRooms.ContainsKey(msg.From.ToString()))
                {
                    if (CScript.Instance.NewDirectMessage(msg.Body, msg.From.ToString(), this))
                    {
                        String chatRoomName = CreateChatRoom(msg.From);
                        CreateChatSession(chatRoomName, msg.From.ToString(), false);
                    }
                }
            }
        }

        private void HandleMessageInChatRoom(Message msg)
        {
            Logger.WriteLine("[MUC-" + _name + "]" + msg.From.Resource + ": " + msg.Body + " To: " + msg.To);
            try
            {
                //Find HTTP Chat session
                ChatSession session = _xmppSessions[msg.From.User];
                //Make sure this message comes from some who is not the Agent
                if (session.AgentName != msg.From.Resource)
                {
                    CScript.Instance.MonitorMessage(msg.Body, msg.From.ToString(), this, false);
                    session.SendChatText(msg.Body);
                }
                else
                {
                    CScript.Instance.MonitorMessage(msg.Body, msg.From.ToString(), this, true);
                }
            }
            catch
            {
                Logger.WriteLine("[" + _name + "] Lost Chat Session:" + msg.From.User);
            }
        }

        #region Utils

        private String CreateChatRoom(Jid from)
        {
            //Create XMPP Chat Room
            String chatRoomName = _name + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            Jid room = CreateRoom(chatRoomName, from.ToString());

            CustomerChat chat = new CustomerChat();
            chat.ChatRoom = room;
            chat.Customer = from;

            _xmppChatRooms.GetOrAdd(from.ToString(), chat);

            return chatRoomName;
        }

        private Boolean CreateChatSession(String chatRoomName, String jID, Boolean sAA)
        {
            //Create HTTP Chat Session
            ChatSession session = new ChatSession();
            session.StartedEvent += Session_StartedEvent;
            session.DisconnectedEvent += Session_DisconnectedEvent;
            session.QueueDataEvent += Session_QueueDataEvent;
            session.SingleAgent = sAA;

            if (session.CreateChatSession(_name, chatRoomName, jID))
            {
                _xmppSessions.GetOrAdd(chatRoomName, session);
                RaiseChatSessionStartedEvent();
                return true;
            }
            else
            {
                Logger.WriteLine("Error Creating Chat Session");
                return false;
            }
        }

        private Boolean CreateChatSession(String jID, Boolean sAA)
        {
            //Create HTTP Chat Session
            ChatSession session = new ChatSession();
            session.StartedEvent += Session_StartedEvent;
            session.DisconnectedEvent += Session_DisconnectedEvent;
            session.SingleAgent = sAA;

            if (session.CreateChatSession(_name, jID))
            {
                _xmppSessions.GetOrAdd(jID, session);
                RaiseChatSessionStartedEvent();

                return true;
            }
            else
            {
                Logger.WriteLine("Error Creating Chat Session");
                return false;
            }
        }

        private Jid CreateRoom(string name, string customerToInvite)
        {
            GroupChatRoom groupChat = new GroupChatRoom();
            groupChat.Chatjid = new Jid(name + "@" + _mUCServerName);
            groupChat.Customerjig = new Jid(customerToInvite);

            _groupChat.AcceptDefaultConfiguration(groupChat.Chatjid, new IqCB(Xmpp_OnCreateRoom), groupChat);

            return groupChat.Chatjid;
        }

        private void RemoveChatSession(String key)
        {
            try
            {
                ChatSession chat;
                _xmppSessions.TryRemove(key, out chat);
                RaiseChatSessionEnded();
            }
            catch { }
        }

        #endregion Utils

        #region Chat Event Handlers

        private void Session_DisconnectedEvent(ChatSession sender)
        {
            Logger.WriteLine("[" + _name + "] Chat Ended");

            CScript.Instance.ChatSessionEnded(sender.XmppCustomerJID.ToString(), this, sender.AgentName);

            if (sender.UseXMPPChatRoom)
            {
                _groupChat.LeaveRoom(new Jid(sender.XmppRoomJID + "@" + _mUCServerName), _name);

                try
                {
                    OutBoundChat chat;
                    _xmppOutBoundSessions.TryRemove(sender.XmppCustomerJID, out chat);
                }
                catch { }

                try
                {
                    CustomerChat chat;
                    _xmppChatRooms.TryRemove(sender.XmppCustomerJID, out chat);
                }
                catch { }

                RemoveChatSession(sender.XmppRoomJID);
            }
            else
            {
                try
                {
                    OutBoundChat chat;
                    _xmppOutBoundSessions.TryRemove(sender.XmppCustomerJID, out chat);
                }
                catch { }
                RemoveChatSession(sender.XmppCustomerJID);
            }
        }

        private void Session_QueueDataEvent(int posInQ, int avgWait)
        {
            Logger.WriteLine("[" + _name + "] avgWait Updated: " + avgWait);
            AvgWaitTime = avgWait;
        }

        private void Session_StartedEvent(ChatSession sender)
        {
            //CScript.Instance.ChatSessionStarted(sender.XmppCustomerJID.ToString(), this, sender.AgentName);
            RaiseChatSessionTakenEvent();

            if (!UseXMPPChatRoom)
            {
                _xmpp.Send(new Message(sender.XmppCustomerJID, MessageType.chat, "Ok, " + sender.AgentName + " is contacting you"));
            }

            Logger.WriteLine("[" + _name + "] Chat Has Started With Agent");
        }

        #endregion Chat Event Handlers

        #region Future...

        public void ShutDown()
        {
            _xmpp.Close();
        }

        #endregion Future...
    }
}