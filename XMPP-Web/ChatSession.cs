using agsXMPP;
using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Timers;

namespace XMPP_Web
{
    internal class ChatSession : IDisposable
    {
        public String SessionState = ChatState.UNKNOWN;

        public Boolean SingleAgent = false;

        internal ChatAgent XmppAgent = null;

        private Boolean _disposed = false;
        private Timer _timer;

        private int Ackcount = -1;

        private String AproposID;

        private int Oldackcount = 0;

        private int Querycount = 0;

        private String SessionID;

        private Object ThreadLock = new Object();

        public String AgentName { private set; get; }

        public String QueueName { private set; get; }

        public DateTime StartTime { private set; get; }

        public Boolean UseXMPPChatRoom { private set; get; }

        public String XmppCustomerJID { private set; get; }

        public String XmppRoomJID { private set; get; }

        public static Boolean Ping()
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://servername/Chat/servlet/AppMain?__lFILE=index.jsp");
            httpWebRequest.Method = "GET";
            try
            {
                HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    return true;
                }
            }
            catch (WebException ex)
            {
                var stream = ex.Response.GetResponseStream();
                using (var streamReader = new StreamReader(stream))
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        public void Agent_LoggedInEvent(object sender, EventArgs e)
        {
            RaiseStartedEvent();
        }

        public Boolean CreateChatSession(String queueName, String xmppRoomName, String xmppUser)
        {
            QueueName = queueName;
            XmppRoomJID = xmppRoomName;
            XmppCustomerJID = xmppUser;
            StartTime = DateTime.UtcNow;
            UseXMPPChatRoom = true;

            if (GetAproposID())
            {
                DoWork();
                return true;
            }
            else
            {
                return false;
            }
        }

        public Boolean CreateChatSession(String queueName, String xmppUser)
        {
            QueueName = queueName;
            XmppRoomJID = "";
            XmppCustomerJID = xmppUser;
            StartTime = DateTime.UtcNow;
            UseXMPPChatRoom = false;

            if (GetAproposID())
            {
                DoWork();
                return true;
            }
            else
            {
                return false;
            }
        }

        public void DisconnectChat()
        {
            if (SessionState == ChatState.CHATTING || SessionState == ChatState.CONNECTED)
            {
                SendClientRequest(++Querycount, "CHAT_STATUS%7CMODE%3DDISCONNECTED");
            }

            if (XmppAgent != null)
            {
                XmppAgent.Disconnect();
            }
            SessionState = ChatState.DISCONNECTED;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void DoWork()
        {
            StartChatSession();

            _timer = new Timer(2000);
            _timer.Enabled = true;
            _timer.AutoReset = true;
            _timer.Elapsed += new ElapsedEventHandler(OnTimedEvent);
        }

        public void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            SendClientRequest(Querycount, "");
        }

        public void SendChatText(String text)
        {
            String phrase = "!$!" + text;
            String query = "CHAT_DATA%7CTEXT%3D" + phrase + "%7CSOURCE%3D1";
            SendClientRequest(++Querycount, query);
        }

        public void SendTypingIndicator(Boolean isTyping)
        {
            if (isTyping)
            {
                SendClientRequest(++Querycount, "TYPING_INDICATOR%7CIS_TYPING%3Dtrue");
            }

            SendClientRequest(++Querycount, "TYPING_INDICATOR%7CIS_TYPING%3Dfalse");
        }

        public void ShutDown()
        {
            _timer.Enabled = false;
            _timer.AutoReset = false;
            XmppAgent.Disconnect();
        }

        private void CreateAgentSession()
        {
            if (UseXMPPChatRoom)
            {
                XmppAgent = new ChatAgent(AgentName, "password", XmppRoomJID);
            }
            else
            {
                XmppAgent = new ChatAgent(AgentName, "password", new Jid(XmppCustomerJID));
                XmppAgent.MessageEvent += XmppAgent_MessageEvent;
            }
            XmppAgent.LoggedInEvent += Agent_LoggedInEvent;
            XmppAgent.Connect();
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    if (_timer != null)
                    {
                        _timer.Dispose();
                    }
                }

                _disposed = true;
            }
        }

        #region Events

        public delegate void DisconnectedEventHandler(ChatSession sender);

        public delegate void QueueDataEventHandler(int posInQ, int avgWait);

        public delegate void StartedEventHandler(ChatSession sender);

        public event DisconnectedEventHandler DisconnectedEvent;

        public event QueueDataEventHandler QueueDataEvent;

        public event StartedEventHandler StartedEvent;

        private void RaiseDisconnectedEvent()
        {
            DisconnectedEventHandler tempEvent = DisconnectedEvent;

            if (tempEvent != null)
            {
                tempEvent(this);
            }
        }

        private void RaiseQueueData(int posInQ, int avgWait)
        {
            QueueDataEventHandler tempEvent = QueueDataEvent;

            if (tempEvent != null)
            {
                tempEvent(posInQ, avgWait);
            }
        }

        private void RaiseStartedEvent()
        {
            StartedEventHandler tempEvent = StartedEvent;

            if (tempEvent != null)
            {
                tempEvent(this);
            }
        }

        #endregion Events

        private String GetAgentName(String command)
        {
            String reg = @"handleTypingIndicator\(true, '([A-Za-z0-9]+)'";
            Match match = Regex.Match(command, reg);

            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                return "Agent";
            }
        }

        private Boolean GetAproposID()
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://servername/Chat/servlet/AppMain?__lFILE=index.jsp");
            httpWebRequest.Method = "GET";
            try
            {
                HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                for (int i = 0; i < httpResponse.Headers.Count; ++i)
                {
                    string header = httpResponse.Headers.GetKey(i);
                    if (header == "Set-Cookie")
                    {
                        foreach (string value in httpResponse.Headers.GetValues(i))
                        {
                            if (value.Contains("JSESSIONID"))
                            {
                                String[] results = value.Split('=');
                                SessionID = results[1];
                            }
                        }
                    }
                }

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    String reg = @"<a href=""/Chat/servlet/AppMain\?__lFILE=OneClick\.jsp&APROPOSID=([0-9a-z]+)"">";
                    var responseText = streamReader.ReadToEnd();
                    Match match = Regex.Match(responseText, reg);
                    AproposID = match.Groups[1].Value;

                    return true;
                }
            }
            catch (WebException ex)
            {
                var stream = ex.Response.GetResponseStream();
                using (var streamReader = new StreamReader(stream))
                {
                    var responseText = streamReader.ReadToEnd();

                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private void HandleClientResult(String result)
        {
            if (result != null)
            {
                String reg = @"var commands = \[(.*)\];";
                var responseText = result;
                RegexOptions options = RegexOptions.Singleline;
                Match match = Regex.Match(responseText, reg, options);

                if (match.Success)
                {
                    String commandArray = match.Groups[1].Value;

                    String expression = @"\[[0-9]+,""(.*)\],";

                    foreach (Match command in Regex.Matches(commandArray, expression))
                    {
                        String num = @"\[([0-9]+),"".*";
                        Match nummatch = Regex.Match(command.Value, num);

                        Ackcount = Convert.ToInt32(nummatch.Groups[1].Value);
                        if (Querycount <= Ackcount)
                        {
                            Querycount = Ackcount;
                        }

                        if (command.Value.Contains("handleTypingIndicator"))
                        {
                            if (XmppAgent == null)
                            {
                                AgentName = GetAgentName(command.Value);
                                CreateAgentSession();
                            }
                            OnTypingIndicator();
                        }
                        else if (command.Value.Contains("handleChatPositionInQueue"))
                        {
                            OnPosInQ(command.Value);
                        }
                        else if (command.Value.Contains("handleChatStatus"))
                        {
                            OnChatStatus(command.Value);
                        }
                        else if (command.Value.Contains("handleChatOutput"))
                        {
                            OnChatOutput(command.Value);
                        }
                    }
                }
            }
        }

        private String HttpPost(String uri, String body)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers.Add("Cookie", "JSESSIONID=" + SessionID);
            httpWebRequest.ContentType = "application/x-www-form-urlencoded";
            httpWebRequest.Headers.Add("Cache-Control", "no-cache");
            httpWebRequest.Referer = "http://servername/Chat/UI/lib/ControlPanel.jsp&APROPOSID=" + AproposID;

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(body);
            }

            try
            {
                HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();

                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var responseText = streamReader.ReadToEnd();
                    return responseText;
                }
            }
            catch (WebException ex)
            {
                var stream = ex.Response.GetResponseStream();
                using (var streamReader = new StreamReader(stream))
                {
                    var responseText = streamReader.ReadToEnd();

                    return null;
                }
            }
        }

        private void OnChatOutput(String command)
        {
            if (command.Contains("SOURCE_AGENT"))
            {
                String reg = @"handleChatOutput\('!\$!(.*)', 'SOURCE_AGENT', .+";
                Match match = Regex.Match(command, reg);

                if (match.Success)
                {
                    String message = match.Groups[1].Value;
                    XmppAgent.SendMessage(message);
                }
            }
        }

        private void OnChatStatus(String command)
        {
            String reg = @".*handleChatStatus\('([A-Z]+)'\);""";
            Match match = Regex.Match(command, reg);
            String status = match.Groups[1].Value;

            if (SessionState == ChatState.UNKNOWN && status == ChatState.WAITING)
            {
                SessionState = ChatState.CONNECTED;
                SendClientRequest(Querycount, "CLIENT_READY");
            }
            else if (status == ChatState.CHATTING && SessionState == ChatState.CONNECTED)
            {
                SessionState = ChatState.CHATTING;
            }
            else if (status == ChatState.DISCONNECTED)
            {
                if (XmppAgent != null)
                {
                    XmppAgent.Disconnect();
                }
                //We may have been disconnect by ChatQueue.cs...dont tell them they already know.
                if (SessionState == ChatState.DISCONNECTED)
                {
                    RaiseDisconnectedEvent();
                }
                else
                {
                    SessionState = ChatState.DISCONNECTED;
                }
            }
        }

        private void OnPosInQ(String command)
        {
            String reg = @"handleChatPositionInQueue\('([0-9]+)', '([0-9]+)'\)";
            Match match = Regex.Match(command, reg);

            if (match.Success)
            {
                try
                {
                    RaiseQueueData(Convert.ToInt32(match.Groups[1].Value), Convert.ToInt32(match.Groups[2].Value));
                }
                catch { }
            }
        }

        private void OnTypingIndicator()
        {
        }

        private void SendClientRequest(int queryNo, String queryData)
        {
            //Lets only send Chat one request at a time in order
            lock (ThreadLock)
            {
                String url = "http://servername/Chat/servlet/com.apropos.weblet.server.ClientReader?APROPOSID=" + AproposID;

                String query = "";
                if (queryNo != -1 && queryData.Length > 0)
                {
                    query = queryNo + "%7C" + queryData + "%0D%0A";
                }

                int ackNo = Ackcount;
                if (Ackcount != Oldackcount)
                {
                    ackNo = Ackcount;
                    Oldackcount = Ackcount;
                }

                String body = "QUERY=" + query +
                        "-1%7CACK%7CACKNO%3D" + ackNo +
                        "%0D%0A&ORIGINATING_HTML0=&ORIGINATING_HTML0URL=&ORIGINATING_HTML1=&ORIGINATING_HTML1URL=&ORIGINATING_HTML2=&ORIGINATING_HTML2URL=&ORIGINATING_HTML3=&ORIGINATING_HTML3URL=&ORIGINATING_HTML4=&ORIGINATING_HTML4URL=";

                String result = HttpPost(url, body);
                HandleClientResult(result);
            }
        }

        private void StartChatSession()
        {
            String sAA = "FALSE";
            if (SingleAgent)
                sAA = "TRUE";

            String url = "http://servername/Chat/servlet/AppMain?__lCMD=newchat&NAME=Steve&REASON_TAG=" + QueueName + "&JID=" + new Jid(XmppCustomerJID).Bare + "&APROPOSID=" + AproposID + "&SAA=" + sAA + "&__lSCRIPT=CheckSAA";
            HttpPost(url, "");
        }

        private void XmppAgent_MessageEvent(object sender, string message)
        {
            SendChatText(message);
        }
    }

    internal class ChatState
    {
        public static String CHATTING { get { return "CHATTING"; } }

        public static String CONNECTED { get { return "CONNECTED"; } }

        public static String DISCONNECTED { get { return "DISCONNECTED"; } }

        public static String UNKNOWN { get { return "UNKNOWN"; } }

        public static String WAITING { get { return "WAITING"; } }
    }
}