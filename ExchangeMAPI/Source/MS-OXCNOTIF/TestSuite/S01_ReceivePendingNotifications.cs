namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Net.NetworkInformation;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test cases for S01_ReceivePendingNotifications.
    /// </summary>
    [TestClass]
    public class S01_ReceivePendingNotifications : TestSuiteBase
    {
        #region Priviate Field

        /// <summary>
        /// The opaque context data used in EcRRegisterPushNotification.
        /// </summary>
        private const string OpaqueContextData = "Opaque data";

        #endregion

        #region Class Initialization and Cleanup

        /// <summary>
        /// Class initialize.
        /// </summary>
        /// <param name="testContext">The session context handle</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is designed to verify that, when the server cannot fit all the notification details into the ROP response buffer, the server uses RopPending to notify the client that there are pending notifications on the server for the client.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC01_VerifyRopPending()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            #region Subscribe ObjectCreated event on server.
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            #endregion

            #region Trigger 200 ObjectCreated events to fill the response buffer fully.

            for (int i = 0; i < 200; ++i)
            {
                this.TriggerObjectCreatedEvent();
            }
            #endregion

            #region Get notification details and pending notification.
            IList<IDeserializable> response = this.CNOTIFAdapter.GetNotification(true);
            #endregion

            #region Verify that after a RopPending ROP response there are additional notifications available on the implementation.

            if (Common.IsRequirementEnabled(81001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R81001: the last response buffer for notification is {0}", response.Last().GetType().Name);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R81001
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "RopPendingResponse",
                    response.Last().GetType().Name,
                    81001,
                    @"[In EcDoRpcExt] [When the value of pcbOut is 0xC350, ]The RopPending ROP ([MS-OXCROPS] section 2.2.14.3) notifies the client that there are pending notifications on the server for the client. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(81002, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R81002: the last response buffer for notification is {0}", response.Last().GetType().Name);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R81002 
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "RopPendingResponse",
                    response.Last().GetType().Name,
                    81002,
                    @"[In EcDoRpcExt] [When the value of pcbOut is 0x190, ]The RopPending ROP ([MS-OXCROPS] section 2.2.14.3) notifies the client that there are pending notifications on the server for the client. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(8201001, this.Site))
            {
                if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site) == "mapi_http")
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R8201001");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R8201001
                    this.Site.CaptureRequirementIfAreEqual<string>(
                        "RopPendingResponse",
                        response.Last().GetType().Name,
                        8201001,
                        @"[In Appendix A: Product Behavior] This ROP RopPending does appear in response buffers of the Execute request type.<7> Section 2.2.1.3.4: The Execute request type was introduced in Exchange 2013 SP1. (Exchange 2013SP1 and above follow this behavior).");
                }
            }
            #endregion

            #region Verify that the pending notification is returned when the response buffer is full.
            bool isNotifyResponse = false;
            for (int i = 0; i < (response.Count - 1); i++)
            {
                if (response[i] is RopNotifyResponse)
                {
                    RopNotifyResponse notifyResponse = (RopNotifyResponse)response[i];
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R97001");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R97001
                    this.Site.CaptureRequirementIfAreEqual<byte>(
                        0x00,
                        notifyResponse.LogonId,
                        97001,
                        @"[In RopNotify ROP Response Buffer] [LogonId] An unsigned integer that specifies the logon associated with the notification event.");

                    isNotifyResponse = true;
                }
                else
                {
                    Site.Assert.IsInstanceOfType(response[i], typeof(RopNotifyResponse), "the implementation does include as many RopNotify ROP responses as will fit in the response before RopPending");
                }
            }

            Site.Assert.IsTrue(isNotifyResponse, "The server should respond a RopNotify Response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R241");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R241
            // The server responds a RopNotify Response, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                241,
                @"[In Sending Notification Details] The RopNotify command is the only method to transmit notification details to the client, so it is used regardless of the method used to notify the client of the pending notification.");

            if (Common.IsRequirementEnabled(348, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R348: the first and last response buffer for notification separately are {0},{1}", response.First().GetType().Name, response.Last().GetType().Name);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R348 
                // Implementation does send as many RopNotify ROP responses as the response buffer allows means if all the RopNotify ROP responses do not fit in the response buffer, the implementation does include as many RopNotify ROP responses as will fit in the response,
                // and then include a RopPending ROP response to indicate that additional notifications are available on the implementation.
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "RopPendingResponse",
                    response.Last().GetType().Name,
                    348,
                    @"[In Appendix A: Product Behavior] Implementation does send as many RopNotify ROP responses as the response buffer allows. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(335, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R335: the first and last response buffer for notification separately are {0},{1}", response.First().GetType().Name, response.Last().GetType().Name);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R335 
                this.Site.CaptureRequirementIfAreEqual<string>(
                    "RopPendingResponse",
                    response.Last().GetType().Name,
                    335,
                    @"[In Appendix A: Product Behavior] If all the RopNotify ROP responses do not fit in the response buffer, the implementation does include as many RopNotify ROP responses as will fit in the response, and then include a RopPending ROP response ([MS-OXCROPS] section 2.2.14.3) to indicate that additional notifications are available on the implementation. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Get all the left notification details.
            // The retry times to try getting all left notifications.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            List<List<uint>> responseSOHs;
            do
            {
                response = this.CNOTIFAdapter.Process(
                        null,
                        this.CNOTIFAdapter.LogonHandle,
                        out responseSOHs);
                retryCount--;
            }
            while (response.Count != 0 && retryCount > 0);
            Site.Assert.IsTrue(
                response.Count == 0,
                "The left notifications aren't all received in {0} retry times. Try to configure RetryCount property in configure file.",
                Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that, when there are pending notifications on the session context associated with the client, the server also sends RopPending to any linked session contexts.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC02_VerifyRopPendingForLinkedSession()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(342, this.Site), "This case runs only under Exchange 2007, since Exchange 2010,Exchange 2013 and Exchange 2016 do not support Session Context linking.");
       
            #region Subscribe ObjectCreated event on server using receiver context handle.
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            #endregion

            #region Client connects to server and logs on the linked session.
            bool isLinkedSessionConnected = this.CNOTIFAdapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.Site.Assert.IsTrue(isLinkedSessionConnected, "If is connected, the ReturnValue of its response is true(success)");
            IntPtr rpcContextForLinkedSession = this.CNOTIFAdapter.RPCContext;

            this.CNOTIFAdapter.Logon();
            uint logonHandleForLinkedSession = this.CNOTIFAdapter.LogonHandle;
            #endregion

            #region Subscribe ObjectCreated event on server using the linked session context handle.
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            #endregion

            #region Trigger many server event using sender context handle to fill the response buffer fully.

            this.CNOTIFAdapter.RPCContext = this.RpcContextForReceive;
            this.CNOTIFAdapter.LogonHandle = this.ReceiverContextLogonHandle;

            // Trigger 200 ObjectCreated events to fill the response buffer fully.
            for (int i = 0; i < 200; ++i)
            {
                this.TriggerObjectCreatedEvent();
            }

            #endregion

            #region Get notification details and a pending notification using receiver context handle.
            IList<IDeserializable> response = this.CNOTIFAdapter.GetNotification(true);
            RopNotifyResponse notifyResponse = (RopNotifyResponse)response.First(x => x is RopNotifyResponse);
            Site.Assert.AreEqual<NotificationType>(NotificationType.ObjectCreated, notifyResponse.NotificationData.NotificationType, "ObjectCreated event should be returned successfully.");
            Site.Assert.AreEqual<string>("RopPendingResponse", response.Last().GetType().Name, "There should be a RopPending response.");

            #endregion

            #region Get all the left notification details on the receiver context handle.
            // The retry times to try getting all left notifications.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            List<List<uint>> responseSOHs;
            do
            {
                response = this.CNOTIFAdapter.Process(
                        null,
                        this.CNOTIFAdapter.LogonHandle,
                        out responseSOHs);
                retryCount--;
            }
            while (response.Count != 0 && retryCount > 0);
            Site.Assert.IsTrue(
                response.Count == 0,
                "The left notifications aren't all received in {0} retry times. Try to configure RetryCount property in configure file.",
                Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            #endregion

            #region Get pending notification and notification details using the linked session context handle.
            this.CNOTIFAdapter.RPCContext = rpcContextForLinkedSession;
            this.CNOTIFAdapter.LogonHandle = logonHandleForLinkedSession;
            IList<IDeserializable> response2 = this.CNOTIFAdapter.GetNotification(true);
            RopNotifyResponse notifyResponse2 = (RopNotifyResponse)response2.First(x => x is RopNotifyResponse);
            Site.Assert.AreEqual<NotificationType>(NotificationType.ObjectCreated, notifyResponse2.NotificationData.NotificationType, "ObjectCreated event should be returned successfully.");

            #endregion

            #region Verify that the pending notification is returned on the linked session context handle.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R342: the last response buffer for notification is {0}", response2.Last().GetType().Name);

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R342 
            this.Site.CaptureRequirementIfAreEqual<string>(
                "RopPendingResponse",
                response2.Last().GetType().Name,
                342,
                @"[In Sending a RopPending ROP Response] The server sends a RopPending ROP response to the client whenever there are pending notifications on any linked session contexts.");
            #endregion

            #region Get all the left notification details on the linked session context handle.
            // The retry times to try getting all left notifications.
            retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));

            do
            {
                response = this.CNOTIFAdapter.Process(
                        null,
                        this.CNOTIFAdapter.LogonHandle,
                        out responseSOHs);
                retryCount--;
            }
            while (response.Count != 0 && retryCount > 0);
            Site.Assert.IsTrue(
                response.Count == 0,
                "The left notifications aren't all received in {0} retry times. Try to configure RetryCount property in configure file.",
                Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            #endregion

            #region Disconnect the linked session connection.

            bool isDisconnected = this.CNOTIFAdapter.DoDisconnect();
            this.Site.Assert.IsTrue(isDisconnected, "If is disconnected, the ReturnValue of its response is true(success)");

            // Switch to the receiver connection.
            this.CNOTIFAdapter.RPCContext = this.RpcContextForReceive;
            this.CNOTIFAdapter.LogonHandle = this.ReceiverContextLogonHandle;
            this.CNOTIFAdapter.IsConnected = this.IsReceiverContextConnected;

            #endregion
        }

        /// <summary>
        /// This test case is designed to implement that the server uses Push Notification to inform the client that the notifications for the session are pending on the server. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC03_VerifyPushNotification()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(313, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 do not support EcRRegisterPushNotification.");

            #region Variables
            int port = this.GetValidUDPPort();
            string opaque = S01_ReceivePendingNotifications.OpaqueContextData;
            #endregion

            #region Subscribe NewMail event
            this.CNOTIFAdapter.RegisterNotification(NotificationType.NewMail);
            #endregion

            #region Call EcRRegisterPushNotification with the valid callback address
            uint resultEcRRegister = this.CNOTIFAdapter.EcRRegisterPushNotification(AddressFamily.AF_INET, port, opaque);
            #endregion

            #region Verify that the server supports at a minimum the AF_INET address type for IP support.

            if (Common.IsRequirementEnabled(321, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R321: the return value of EcRRegisterPushNotification is {0}", resultEcRRegister);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R321
                // Use AF_INET address type to register push notification, if a successful response is returned means implementation support this address type.
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00,
                    resultEcRRegister,
                    321,
                     @"[In Appendix A: Product Behavior] Implementation does support at a minimum the AF_INET address type for IP support. (Exchange 2007 follows this behavior.)");
            }

            #region Verify that Exchange 2007 support the EcRRegisterPushNotification method call
            if (Common.IsRequirementEnabled(73, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R73");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R73
                this.Site.CaptureRequirementIfAreEqual<uint>(
                     0x00,
                     resultEcRRegister,
                     73,
                     @"[In Appendix A: Product Behavior] The EcRRegisterPushNotification RPC method, as specified in [MS-OXCRPC] section 3.1.4.5, is used to register a callback address of a client on the implementation. (Exchange 2007 follows this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R313");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R313
            this.Site.CaptureRequirementIfAreEqual<uint>(
                 0x00,
                 resultEcRRegister,
                 313,
                 @"[In Appendix A: Product Behavior] Implementation does support the EcRRegisterPushNotification method call. (Section 3.1.5.4: Exchange 2007 supports push notifications and the EcRRegisterPushNotification method, as specified in [MS-OXCRPC] section 3.1.4.5.)");
            #endregion
            #endregion

            #region Trigger the NewMail event and get the push notification datagram
            this.TriggerNewMailEvent();
            string opaqueReturned;

            bool pushResult = this.CNOTIFAdapter.PushNotificationReceived(AddressFamily.AF_INET, port, out opaqueReturned);
            Site.Assert.IsTrue(pushResult, "The client failed to get UDP package from server. Increase the timeout value defined in PushNotificationTimeout property in ptfconfig file, and try again.");

            if (Common.IsRequirementEnabled(538, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R538");

                pushResult = this.CNOTIFAdapter.PushNotificationReceived(AddressFamily.AF_INET, port, out opaqueReturned);
                DateTime secondTimeGetUDPDatagram = DateTime.Now;
                this.Site.Log.Add(LogEntryKind.Debug, secondTimeGetUDPDatagram.ToString() + " Get the second UDP datagram from server.");
                Site.Assert.IsTrue(pushResult, "The client failed to get UDP package from server. Increase the timeout value defined in PushNotificationTimeout property in ptfconfig file, and try again.");

                pushResult = this.CNOTIFAdapter.PushNotificationReceived(AddressFamily.AF_INET, port, out opaqueReturned);
                DateTime thirdTimeGetUDPDatagram = DateTime.Now;
                this.Site.Log.Add(LogEntryKind.Debug, thirdTimeGetUDPDatagram.ToString() + " Get the third UDP datagram from server.");
                Site.Assert.IsTrue(pushResult, "The client failed to get UDP package from server. Increase the timeout value defined in PushNotificationTimeout property in ptfconfig file, and try again.");

                // The server continue sending a UDP datagram to the callback address.
                // Get the time interval between the second time and the third time.  
                TimeSpan interval = thirdTimeGetUDPDatagram.Subtract(secondTimeGetUDPDatagram);

                int result = interval.CompareTo(new TimeSpan(0, 0, 60));

                // The deviation of serval seconds is acceptable.
                int timeDeviation = int.Parse(Common.GetConfigurationPropertyValue("UDPDatagramsIntervalDeviation", this.Site));
                bool isVerifiedR538 = result < timeDeviation || result > (0 - timeDeviation);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R538
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR538,
                    538,
                    @"[In Appendix A: Product Behavior] Implementation does allow for a 60-second interval between UDP datagrams until the client has retrieved all event information for the session, if push notifications are supported by the implementation. (Exchange 2007 follows this behavior.)");

                if (Common.IsRequirementEnabled(539, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R539");

                    // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R539
                    // MS-OXCNOTIF_R538 is verified, MS-OXCNOTIF_R539 can be verified directly.
                    this.Site.CaptureRequirement(
                        539,
                        @"[In Appendix A: Product Behavior] Implementation does continue sending a UDP datagram to the callback address at 60-second intervals if event details are still queued for the client.(Exchange 2007 follows this behavior.)");
                }
            }

            IList<IDeserializable> response = this.CNOTIFAdapter.GetNotification(true);
            #endregion

            #region Verify that the callback address is used to support push notifications.
            bool isSupportpushNotification = false;
            foreach (IDeserializable resp in response)
            {
                if (resp is RopNotifyResponse)
                {
                    isSupportpushNotification = true;
                }
            }

            Site.Assert.IsTrue(pushResult, "Push notification on the specified port should be received.");
            Site.Assert.IsTrue(isSupportpushNotification, "The server should send a RopNotify ROP response use the callback address.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R79");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R79
            // Callback address has been registered in EcRRegisterPushNotification method, 
            // and the server send a RopNotify ROP response use the callback address, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                79,
                @"[In EcRRegisterPushNotification Method] It [callback address] is used to support push notifications, which is one way in which the server can notify clients of pending notifications.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R80");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R80
            // Use push notification with UDP datagrams and get the notification successfully, this requirement can be verified directly.
            this.Site.CaptureRequirement(
                80,
                @"[In EcRRegisterPushNotification Method] The UDP datagrams inform the client that notifications are pending on the server for the session.");

            #endregion

            #region Verify that the server sends a push notification datagram which contains the client's opaque data

            if (Common.IsRequirementEnabled(326, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R326");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R326
                this.Site.CaptureRequirementIfAreEqual<string>(
                    opaque,
                    opaqueReturned,
                    326,
                    @"[In Appendix A: Product Behavior] After the callback address has been successfully registered with the implementation, the implementation does send a UDP datagram containing the client's opaque data, from the rgbContext field, when a notification becomes available for the client. (Exchange 2007 follows this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R324");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R324
            // The server can successfully send the UDP datagram to the callback address and the same opaque context data
            // as the client sent, so this requirement can be verified.
            bool isVerifiedR324 = pushResult && (opaqueReturned == opaque);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR324,
                324,
                @"[In Receiving an EcRRegisterPushNotification Method Call] The server MUST save the callback address and opaque context data on the session context for future use.");

            #endregion

            #region Retrieve all of the notification details

            // The retry times to try getting all left notifications.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
            List<List<uint>> responseSOHs;
            do
            {
                response = this.CNOTIFAdapter.Process(
                        null,
                        this.CNOTIFAdapter.LogonHandle,
                        out responseSOHs);
                Thread.Sleep(sleepTime);
                retryCount--;
            }
            while (response.Count != 0 && retryCount > 0);
            Site.Assert.IsTrue(
                response.Count == 0,
                "The left notifications aren't all received in {0} retry times. Try to configure RetryCount property in configure file.",
                Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            #endregion

            #region Verify that the server stops sending datagrams when all of the notifications have been retrieved
            // Sleep some time to wait for server sending all left notifications and UDP datagram to client.WaitTime
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            Thread.Sleep(waitTime);
            pushResult = this.CNOTIFAdapter.PushNotificationReceived(AddressFamily.AF_INET, port, out opaqueReturned);

            // When all of the notifications have been retrieved from the implementation, verified whether the Implementation stops sending UDP datagrams.
            if (Common.IsRequirementEnabled(331, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R331: there is {0} notification returned from server", pushResult ? string.Empty : "not");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R331
                bool isVerifiedR331 = !pushResult;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR331,
                    331,
                    @"[In Appendix A: Product Behavior] Implementation does stop sending UDP datagrams when all of the notifications have been retrieved from the implementation through EcDoRpcExt2 method calls, as specified in [MS-OXCRPC] section 3.1.4.2. (Exchange 2007 follows this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that the server supports at a minimum the AF_INET6 address type for IPv6 support. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC04_VerifyPushNotificationForIPv6()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(313, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 do not support EcRRegisterPushNotification.");

            #region Variables
            int port = this.GetValidUDPPort();
            #endregion

            #region Subscribe NewMail event
            this.CNOTIFAdapter.RegisterNotification(NotificationType.NewMail);
            #endregion

            #region Call EcRRegisterPushNotification with the valid IPV6 callback address
            uint resultEcRRegister = this.CNOTIFAdapter.EcRRegisterPushNotification(AddressFamily.AF_INET6, port, S01_ReceivePendingNotifications.OpaqueContextData);
            #endregion

            #region Verify that the server supports at a munimum the AF_INET6 address type for IPv6 support.
            if (Common.IsRequirementEnabled(323, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R323: the return value of EcRRegisterPushNotification is {0}", resultEcRRegister);

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R323
                // Use AF_INET6 address type to register push notification, if a successful response is returned means implementation support this address type.
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00,
                    resultEcRRegister,
                    323,
                    @"[In Appendix A: Product Behavior] Implementation does support at a minimum the AF_INET6 address type for IPv6 support. (Exchange 2007 follows this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that the server fails an EcRRegisterPushNotification call with an invalid callback address.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC05_VerifyPushNotificationFailed()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(313, this.Site), "This case runs only under Exchange 2007, since Exchange 2010 and Exchange 2013 do not support EcRRegisterPushNotification.");

            #region Variables
            int port = this.GetValidUDPPort();
            #endregion

            #region Subscribe NewMail event
            this.CNOTIFAdapter.RegisterNotification(NotificationType.NewMail);
            #endregion

            #region Call EcRRegisterPushNotification with the invalid callback address
            uint resultEcRRegister = this.CNOTIFAdapter.EcRRegisterPushNotification(AddressFamily.Invalid, port, S01_ReceivePendingNotifications.OpaqueContextData);
            #endregion

            #region Verify that the EcRRegisterPushNotification failed
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R318");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R318
            // As specified in MS-OXCRPC section 3.1.4.5, if the EcRRegisterPushNotification method succeeds, the return value is 0. 
            // If the method fails, the return value is an implementation-specific error code or one of the protocol-defined error codes listed in MS-OXCRPC section 3.1.4.5.
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                resultEcRRegister,
                318,
                @"[In Receiving an EcRRegisterPushNotification Method Call] The server MUST fail the call [EcRRegisterPushNotification Method Call], if the callback address is not a valid SOCKADDR structure.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to implement that the server uses Asynchronous RPC Notification to inform the client that notifications are pending on the server for the session.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC06_VerifyAsyncRpcCall()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();

            // Asynchronous RPC Notification can't be verified when use MAPIHTTP as transport.
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() != "mapi_http", "Asynchronous RPC Notification can't be verified when use MAPIHTTP as transport.");

            #region Subscribe NewMail event
            this.CNOTIFAdapter.RegisterNotification(NotificationType.NewMail);
            #endregion

            #region Call EcDoAsyncConnectEx to acquire an asynchronous context handle
            IntPtr acxh = this.CNOTIFAdapter.EcDoAsyncConnectEx();
            #endregion

            #region Verify that the server created an asynchronous context handle
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R298");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R298
            this.Site.CaptureRequirementIfAreNotEqual<IntPtr>(
                IntPtr.Zero,
                acxh,
                298,
                @"[In Receiving an EcDoAsyncConnectEx Method Call] When a call to the EcDoAsyncConnectEx RPC, as specified in [MS-OXCRPC] section 3.1.4.4, is received by the server, the server MUST create an asynchronous context handle.");

            if (Common.IsRequirementEnabled(297, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R297");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R297
                this.Site.CaptureRequirementIfAreNotEqual<IntPtr>(
                    IntPtr.Zero,
                    acxh,
                    297,
                    @"[In Appendix A: Product Behavior] Implementation does support the EcDoAsyncConnectEx method call, as specified in [MS-OXCRPC] section 3.1.4.4. ( Exchange 2007 and above follow this behavior).");
            }

            #endregion

            #region Call EcDoAsyncWaitEx with the valid ACXH
            IntPtr asyncHandle;
            this.CNOTIFAdapter.BeginAsyncWait(acxh, out asyncHandle);

            // Get the status of EcDoAsyncWaitEx call.
            // The event has not been triggered yet, so the call should not be completed.
            RPCAsyncStatus status = this.CNOTIFAdapter.QueryAsyncWaitStatus(asyncHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R62");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R62
            // When the client can get a valid handle and the handle can be used in BeginAsyncWait successfully, this requirement can be verified.
            bool isVerifiedR62 = acxh != IntPtr.Zero && status != RPCAsyncStatus.RPC_S_INVALID_ASYNC_HANDLE;

            if (Common.IsRequirementEnabled(62, this.Site))
            {
                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR62,
                    62,
                    @"[In Appendix A: Product Behavior] The EcDoAsyncConnectEx RPC method, as specified in [MS-OXCRPC] section 3.1.4.4, is used to acquire an asynchronous context handle on the implementation to use in subsequent EcDoAsyncWaitEx method calls, as specified in [MS-OXCRPC] section 3.3.4.1.  (Exchange 2007 and above follow this behavior.)");
            }

            // Trigger the event
            this.TriggerNewMailEvent();

            // Get the status of EcDoAsyncWaitEx call again.
            // The event has been triggered, so the call should be completed.
            RPCAsyncStatus status2;

            // The times to try getting a completed status.
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
            do
            {
                status2 = this.CNOTIFAdapter.QueryAsyncWaitStatus(asyncHandle);
                Thread.Sleep(sleepTime);
                retryCount--;
            }
            while (status2 != RPCAsyncStatus.RPC_S_OK && retryCount >= 0);
            Site.Assert.AreEqual<RPCAsyncStatus>(RPCAsyncStatus.RPC_S_OK, status2, "RPC status should be completed.");
            #endregion

            #region Verify that the server does not complete the call until there is a notification
            if (Common.IsRequirementEnabled(305, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R305");

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R305
                bool isR305Satisfied = (status != RPCAsyncStatus.RPC_S_OK) && (status2 == RPCAsyncStatus.RPC_S_OK);

                this.Site.CaptureRequirementIfIsTrue(
                    isR305Satisfied,
                    305,
                    @"[In Appendix A: Product Behavior] Implementation does not complete the call [EcDoAsyncWaitEx Method Call] until there is a notification for the client session. (Exchange 2007 and above follow this behavior.)");
            }

            #endregion

            #region Complete the EcDoAsyncWaitEx call
            int isPending;
            this.CNOTIFAdapter.EndAsyncWait(asyncHandle, out isPending);
            #endregion

            #region Verify that the server returns the value NotificationPending in the output field pulFlagsOut.

            if (Common.IsRequirementEnabled(302, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R302");

                // The RPC method BeginAsyncWait and EndAsyncWait can run successfully without exception, so this requirement can be verified directly.
                Site.CaptureRequirement(
                302,
                @"[In Appendix A: Product Behavior] Implementation does support the EcDoAsyncWaitEx method call, as specified in [MS-OXCRPC] section 3.3.4.1. (Exchange 2007 and above follow this behavior).");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R310");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R310
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                isPending,
                310,
                @"[In Receiving an EcDoAsyncWaitEx Method Call] If the server completes the outstanding RPC call when there is a notification for the client session, the server MUST return the value NotificationPending in the pulFlagsOut field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R68");

            if (Common.IsRequirementEnabled(68, this.Site))
            {
                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R68
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000001,
                    isPending,
                    68,
                    @"[In Appendix A: Product Behavior] The EcDoAsyncWaitEx asynchronous RPC method, as specified in [MS-OXCRPC] section 3.3.4.1, is used to inform a client about pending notifications on the implementation. ( Exchange 2007 and above follow this behavior)");
            }

            #endregion

            #region Get notification details.
            IList<IDeserializable> response = this.CNOTIFAdapter.GetNotification(true);
            RopNotifyResponse notifyResponse = (RopNotifyResponse)response.First(x => x is RopNotifyResponse);
            Site.Assert.AreEqual<NotificationType>(NotificationType.NewMail, notifyResponse.NotificationData.NotificationType, "New mail should be returned successfully.");

            bool isSupportRpcNotif = false;
            foreach (IDeserializable resp in response)
            {
                if (resp is RopNotifyResponse)
                {
                    isSupportRpcNotif = true;
                }
            }

            #region Verify that the EcDoAsyncConnectEx and EcDoAsyncConnectEx support asynchronous RPC notifications
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R63");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R63
            // EcDoAsyncConnectEx has been called before getting notification, so if the response contains RopNotifyResponse, this requirement can be verified.
            bool isVerifiedR63 = isSupportRpcNotif;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR63,
                63,
                @"[In EcDoAsyncConnectEx Method] The EcDoAsyncConnectEx method is used to support asynchronous RPC notifications.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R69");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R69
            // EcDoAsyncWaitEx has been called before getting notification, so if the response contains RopNotifyResponse, this requirement can be verified.
            bool isVerifiedR69 = isSupportRpcNotif;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR69,
                69,
                @"[In EcDoAsyncWaitEx Method] The EcDoAsyncWaitEx method is used to support asynchronous RPC notifications.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that the server fails an EcDoAsyncWaitEx call with an invalid asynchronous context handle.
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC07_VerifyAsyncRpcCallFailed()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() != "mapi_http", "EcDoAsyncWaitEx call can't be verified when use MAPIHTTP as transport.");

            #region Subscribe NewMail event
            this.CNOTIFAdapter.RegisterNotification(NotificationType.NewMail);
            #endregion

            #region Call EcDoAsyncWaitEx with a invalid ACXH
            IntPtr asyncHandle;
            SEHException exception = null;
            try
            {
                this.CNOTIFAdapter.BeginAsyncWait(IntPtr.Zero, out asyncHandle);
            }
            catch (SEHException e)
            {
                exception = e;
            }
            #endregion

            #region Verify that the EcDoAsyncConnectEx failed
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R303");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R303
            // If there is an exception is thrown when an asynchronous EcDoAsyncWaitEx method call with a invalid ACXH is received, indicate server will validate the asynchronous context handle.
            this.Site.CaptureRequirementIfIsNotNull(
                exception,
                303,
                @"[In Receiving an EcDoAsyncWaitEx Method Call] Whenever an asynchronous EcDoAsyncWaitEx method call, as specified in [MS-OXCRPC] section 3.3.4.1, on the AsyncEMSMDB interface is received by the server, the server MUST validate that the asynchronous context handle provided is a valid asynchronous context handle that was returned from the EcDoAsyncConnectEx method call, as specified in [MS-OXCRPC] section 3.1.4.4.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R434");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R434
            this.Site.CaptureRequirementIfIsNotNull(
                exception,
                434,
                @"[In Sending and Receiving EcDoAsyncWaitEx Method Calls] If the EcDoAsyncWaitEx method returns a non-zero result code, it indicates that an error occurred.");

            #endregion
        }

        /// <summary>
        /// This test case is designed to verify that the server doesn't complete the EcDoAsyncWaitEx call until the call has been outstanding on the server for 5 minutes. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC08_VerifyAsyncRpcCallTimeOut()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() != "mapi_http", "EcDoAsyncWaitEx call can't be verified when use MAPIHTTP as transport.");

            #region Subscribe NewMail event
            uint notificationHandle;
            this.CNOTIFAdapter.RegisterNotificationWithParameter(NotificationType.ObjectMoved, 0, this.NewFolderId, this.TriggerMessageId, out notificationHandle);
            #endregion

            #region Call EcDoAsyncConnectEx to acquire an asynchronous context handle
            IntPtr acxh = this.CNOTIFAdapter.EcDoAsyncConnectEx();
            #endregion

            #region Call EcDoAsyncWaitEx without triggering the event
            IntPtr asyncHandle;
            this.CNOTIFAdapter.BeginAsyncWait(acxh, out asyncHandle);

            // Wait several minutes until the call time out,the time is implementation-specific, which can be configured.
            TimeSpan maxWaitTime = TimeSpan.FromMinutes(int.Parse(Common.GetConfigurationPropertyValue("MaxWaitTime", this.Site)));
            DateTime beginTime = DateTime.Now;
            DateTime endTime;
            bool isCompleted;
            int sleepTime = int.Parse(Common.GetConfigurationPropertyValue("SleepTime", this.Site));
            do
            {
                RPCAsyncStatus status = this.CNOTIFAdapter.QueryAsyncWaitStatus(asyncHandle);
                isCompleted = status == RPCAsyncStatus.RPC_S_OK;
                Site.Log.Add(LogEntryKind.Comment, "The status of EcDoAsyncWaitEx is {0}", status);
                endTime = DateTime.Now;
                Thread.Sleep(sleepTime);
            }
            while (!isCompleted && (endTime - beginTime) < maxWaitTime);

            TimeSpan actualTimeout = endTime - beginTime;

            int isPending;
            this.CNOTIFAdapter.EndAsyncWait(asyncHandle, out isPending);
            #endregion

            #region Verify that the server does not complete the call until the call has been outstanding on the server 5 minutes.
            if (Common.IsRequirementEnabled(307, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R307: The actual time out is {0}", actualTimeout.TotalMinutes);

                // Considering the network, test case run consume, there might be some deviation between the actual and expected time out. 
                // The deviation value is implementation-specific, which can be configured.
                bool timeOutValid = Math.Abs(actualTimeout.TotalMinutes - 5) < int.Parse(Common.GetConfigurationPropertyValue("AsyncWaitTimeoutDeviation", this.Site));

                // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R307
                bool isR307Satisfied = isCompleted && timeOutValid;

                this.Site.CaptureRequirementIfIsTrue(
                    isR307Satisfied,
                    307,
                    @"[In Appendix A: Product Behavior] Implementation does not complete the call [EcDoAsyncWaitEx Method Call] until the call has been outstanding on the server 5 minutes. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion

            #region Verify that the server return 0 in pulFlagsOut if the call was completed when there is no notification for the client session.
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R311");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R311
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                isPending,
                311,
                @"[In Receiving an EcDoAsyncWaitEx Method Call] The server MUST return 0x00000000 in the pulFlagsOut field if the call [outstanding RPC call] was completed when there is no notification for the client session.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to verify the NotificationWait request which is used to notify the client of pending notifications. 
        /// </summary>
        [TestCategory("MSOXCNOTIF"), TestMethod()]
        public void MSOXCNOTIF_S01_TC09_VerifyNotificationWait()
        {
            this.CheckWhetherSupportMAPIHTTP();
            this.NotificationInitialize();
            Site.Assume.IsTrue(Common.IsRequirementEnabled(482, this.Site) && Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http", "This case runs only under Exchange 2013 SP1 using MAPIHTTP as transport, since Exchange 2007 and Exchange 2010 do not support MAPIHTTP, and Exchange 2013 SP1 use RPC as transport does not support NotificationWait request type.");

            #region Subscribe ObjectCreated event on server.
            this.CNOTIFAdapter.RegisterNotification(NotificationType.ObjectCreated);
            #endregion

            #region Trigger ObjectCreated event to fill the response buffer fully.
            this.TriggerObjectCreatedEvent();
            #endregion

            #region Send NotificationWait request to request that the server notify the client of pending notifications.
            NotificationWaitRequestBody notificationWaitRequestBody = new NotificationWaitRequestBody()
            {
                Flags = 0,
                AuxiliaryBuffer = new byte[] { },
                AuxiliaryBufferSize = 0
            };

            NotificationWaitSuccessResponseBody notificationWaitResponseBody = this.CNOTIFAdapter.NotificationWait(notificationWaitRequestBody);

            this.Site.Assert.AreEqual<uint>(
                0,
                notificationWaitResponseBody.ErrorCode,
                "If NotificationWait request succeeds, the ErrorCode of its response is 0 (success)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCNOTIF_R482");

            // Verify MS-OXCNOTIF requirement: MS-OXCNOTIF_R482
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                notificationWaitResponseBody.EventPending,
                482,
                "[In Appendix A: Product Behavior] Implementation uses NotificationWait request type to notify the client of pending notifications. (Exchange 2013 SP1 and above follow this behavior.)");
            #endregion
        }
        #endregion

        /// <summary>
        /// Get the valid UDP port on local machine. The default port number is 1025.
        /// </summary>
        /// <returns>Return the valid port number on local machine.</returns>
        private int GetValidUDPPort()
        {
            IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();
            IPEndPoint[] udpEndpoints = properties.GetActiveUdpListeners();
            IPEndPoint[] tcpEndpoints = properties.GetActiveTcpListeners();

            int portNumber = int.Parse(Common.GetConfigurationPropertyValue("NotificationPort", this.Site));
            int newPort = portNumber;

            // To check if the portNumber is used by other applications (both UDP and TCP).
            // After the loop, get the valid port number newPort, which will be used in the test case. 
            do
            {
                portNumber = newPort;
                foreach (IPEndPoint point in udpEndpoints)
                {
                    if (point.Port == portNumber)
                    {
                        newPort++;
                        break;
                    }
                }

                if (newPort == portNumber)
                {
                    foreach (IPEndPoint point in tcpEndpoints)
                    {
                        if (point.Port == portNumber)
                        {
                            newPort++;
                            break;
                        }
                    }
                }
            }
            while (newPort != portNumber);

            Site.Log.Add(LogEntryKind.Debug, "The valid UDP port is {0}", newPort);
            return newPort;
        }
    }
}