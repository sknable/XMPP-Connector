using System;
using XMPP_Web;
using System.Threading;

    public class XMPP_Script
    {
        public void Initialize()
        {
            
        }

        //Warning ChatQueue can be null!
        public void OnMonitorMessage(String message,String customerJID,ChatQueue queue,Boolean isAgent)
        {


        }

        public void OnChatInQueueUpdate(int numInQueue,ChatQueue queue)
        {
            if(queue.Name == "A-support")
            {
                if (numInQueue >= 1)
                {
                    queue.UpdatePresence("We are experiencing long waiting times", false);
                }
                else
                {
                    queue.UpdatePresence(numInQueue + " Waiting in Queue", true);
                }
            }
            else
            { 
                queue.UpdatePresence(numInQueue + " Waiting in Queue", true);
            }
        }


        public Boolean OnNewDirectMessage(String message, String from, ChatQueue queue)
        {
            Thread.Sleep(500);

            if (queue.UseXMPPChatRoom)
            {
                queue.SendChatMessageAsQueue("Welcome to " + queue.Name + ".", from);
                Thread.Sleep(300);
                queue.SendChatMessageAsQueue("Looks like my average wait time is around " + queue.AvgWaitTime + " seconds.", from);
                Thread.Sleep(700);
                queue.SendChatMessageAsQueue("I'm sending an invite", from);

                if (queue.Name.ToLower() == "A-support")
                {

                    String html = "";

                    queue.SendChatMessageAsQueue("", html, from);

                }
                
                
            }
            else
            {
                queue.SendChatMessageAsQueue("Welcome to " + queue.Name + ". Let me find an Agent for you", from);
            }

            return true;

        }

        public Boolean OnOutBoundMessageConfirmation(String message, String from, ChatQueue queue)
        {
            Thread.Sleep(500);
            if (queue.UseXMPPChatRoom)
            {
                queue.SendChatMessageAsQueue("Ok, thanks Let me send you a invite", from);
            }
            else
            {
                queue.SendChatMessageAsQueue("Ok, I'll go find them", from);
            }
            return true;
        }

        public void OnChatSessionStarted(String customer, ChatQueue queue, String agentName)
        {


        }


        public void OnChatSessionEnded(String customer, ChatQueue queue, String agentName)
        {


        }
    }


