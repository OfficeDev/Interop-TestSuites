//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;

    /// <summary>
    /// This class is used to perform the discovery process between the shared test case part and the pure WOPI test cases part.
    /// </summary>
    public class DiscoveryProcessHelper : HelperBase
    {   
        /// <summary>
        /// A bool value is used to indicating whether the discovery process has been executed successfully.
        /// </summary>
        private static bool hasPerformDiscoveryProcessSucceed = false;

        /// <summary>
        /// A bool value is used to indicating whether the discovery record has been cleaned up successfully.
        /// </summary>
        private static bool hasPerformCleanUpForDiscovery = false;

        /// <summary>
        /// An object instance is used for lock blocks which is used for multiple threads. This instance is used to keep asynchronous process for visiting the "hasPerformDiscoveryProcessSucceed" field in different threads.   
        /// </summary>
        private static object lockObjectOfVisitDiscoveryProcessStatus = new object();

        /// <summary>
        /// An DiscoveryRequestListener instance represents the listener which is listening discovery request. The test suite will use only one listener instance to listen.
        /// </summary>
        private static DiscoveryRequestListener discoveryListenerInstance = null;

        /// <summary>
        /// A string value represents the progId which is used in discovery process in order to make the WOPI server enable folder level visit ability when it receive the progId from the discovery response. 
        /// </summary>
        private static string progIdValue = string.Empty;

        /// <summary>
        /// Prevents a default instance of the DiscoveryProcessHelper class from being created
        /// </summary>
        private DiscoveryProcessHelper()
        { 
        }

        /// <summary>
        /// Gets a value indicating whether the helper need to perform a clean up action for discovery record.
        /// </summary>
        public static bool NeedToCleanUpDiscoveryRecord
        {
            get
            {
                lock (lockObjectOfVisitDiscoveryProcessStatus)
                {
                    return hasPerformDiscoveryProcessSucceed && !hasPerformCleanUpForDiscovery;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether the discovery process has been executed successfully
        /// </summary>
        public static bool HasPerformDiscoveryProcessSucceed
        {
            get
            {
                lock (lockObjectOfVisitDiscoveryProcessStatus)
                {
                    return hasPerformDiscoveryProcessSucceed;
                }
            }
        }

        /// <summary>
        /// A method is used to perform the WOPI discovery process for the WOPI server.
        /// </summary>
        /// <param name="hostNameOfDiscoveryListener">A parameter represents the machine name which is hosting the discovery listener feature. It should be the test client name which is running the test suite.</param>
        /// <param name="sutControllerInstance">A parameter represents the IMS_WOPISUTControlAdapter instance which is used to make the WOPI server perform sending discovery request to the discovery listener.</param>
        public static void PerformDiscoveryProcess(string hostNameOfDiscoveryListener, IMS_WOPISUTControlAdapter sutControllerInstance)
        {
            if (HasPerformDiscoveryProcessSucceed)
            {
                return;
            }

            if (null == sutControllerInstance)
            {
                throw new ArgumentNullException("sutControllerInstance");
            }

            if (string.IsNullOrEmpty(hostNameOfDiscoveryListener))
            {
                throw new ArgumentNullException("hostNameOfDiscoveryListener");
            }

            // Call the "TriggerWOPIDiscovery" method of IMS_WOPISUTControlAdapter interface
            bool isDiscoverySuccessful = sutControllerInstance.TriggerWOPIDiscovery(hostNameOfDiscoveryListener);
            if (!isDiscoverySuccessful)
            {
                throw new InvalidOperationException("Could not perform the discovery process successfully.");
            }

            lock (lockObjectOfVisitDiscoveryProcessStatus)
            {
                hasPerformDiscoveryProcessSucceed = true;
            }

            DiscoveryProcessHelper.AppendLogs(typeof(DiscoveryProcessHelper), DateTime.Now, "Perform the trigger WOPI discovery process successfully.");
        }

        /// <summary>
        /// A method is used to clean up the WOPI discovery record for the WOPI server. For removing the record successfully, the WOPI server can be triggered the WOPI discovery process again.
        /// </summary>
        /// <param name="wopiClientName">A parameter represents the WOPI client name which should have been discovered by WOPI server</param>
        /// <param name="sutControllerInstance">A parameter represents the IMS_WOPISUTControlAdapter instance which is used to make the WOPI server clean up discovery record for the specified WOPI client.</param>
        public static void CleanUpDiscoveryRecord(string wopiClientName, IMS_WOPISUTControlAdapter sutControllerInstance)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<IMS_WOPISUTControlAdapter>(sutControllerInstance, "sutControllerInstance", "CleanUpDiscoveryRecord");

            if (!NeedToCleanUpDiscoveryRecord)
            {
                return;
            }

            lock (lockObjectOfVisitDiscoveryProcessStatus)
            {
               if (hasPerformDiscoveryProcessSucceed && !hasPerformCleanUpForDiscovery)
               {
                   bool isDiscoveryRecordRemoveSuccessful = sutControllerInstance.RemoveWOPIDiscoveryRecord(wopiClientName);
                   if (!isDiscoveryRecordRemoveSuccessful)
                   {
                       throw new InvalidOperationException("Could not remove the discovery record successfully, need to remove that manually.");
                   }

                   hasPerformCleanUpForDiscovery = true;
               }
            }
        }

        /// <summary>
        /// A method is used to generate response of a WOPI discovery request. It indicates the WOPI client supports 3 types file extensions: ".txt", ".zip", ".one"  
        /// </summary>
        /// <param name="currentTestClientName">A parameter represents the current test client name which is used to construct WOPI client's app name in WOPI discovery response.</param>
        /// <param name="progId">A parameter represents the id of program which is associated with folder level visit in discovery process. This value must be valid for WOPI server.</param>
        /// <returns>A return value represents the response of a WOPI discovery request.</returns>
        public static string GetDiscoveryResponseXmlString(string currentTestClientName, string progId)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(currentTestClientName, "currentTestClientName", "GetDiscoveryResponseXmlString");
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(progId, "progId", "GetDiscoveryResponseXmlString");

            wopidiscovery wopiDiscoveryInstance = new wopidiscovery();

            // Pass the prog id, so that the WOPI discovery response logic will use the prog id value.
            progIdValue = progId;

            // Add http and https net zone into the wopiDiscovery
            wopiDiscoveryInstance.netzone = GetNetZonesForWopiDiscoveryResponse(currentTestClientName);

            // ProofKey element
            wopiDiscoveryInstance.proofkey = new ct_proofkey();
            wopiDiscoveryInstance.proofkey.oldvalue = RSACryptoContext.PublicKeyStringOfOld;
            wopiDiscoveryInstance.proofkey.value = RSACryptoContext.PublicKeyStringOfCurrent;
            string xmlStringOfResponseDiscovery = WOPISerializerHelper.GetDiscoveryXmlFromDiscoveryObject(wopiDiscoveryInstance);

            return xmlStringOfResponseDiscovery;
        }

        /// <summary>
        /// A method is used to start listening the discovery request from the WOPI server. If the listen thread has existed, the DiscoverProcessHelper will not start any new listen thread.
        /// </summary>
        /// <param name="currentTestClient">A parameter represent the current test client which acts as WOPI client to listen the discovery request.</param>
        /// <param name="progId">A parameter represents the id of program which is associated with folder level visit in discovery process. This value must be valid for WOPI server. For Microsoft products, this value can be "OneNote.Notebook". It is used to ensure the WOPI server can support folder level visit ability in WOPI mode when receive the value from the discovery response.</param>
        public static void StartDiscoveryListen(string currentTestClient, string progId)
        {
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(currentTestClient, "currentTestClient", "StartDiscoveryListen");
            DiscoveryProcessHelper.CheckInputParameterNullOrEmpty<string>(progId, "progId", "StartDiscoveryListen");

           if (null == discoveryListenerInstance)
           {
               string discoveryXmlResponse = GetDiscoveryResponseXmlString(currentTestClient, progId);
               discoveryListenerInstance = new DiscoveryRequestListener(currentTestClient, discoveryXmlResponse);
               discoveryListenerInstance.StartListen();
           }
        }

        /// <summary>
        /// A method is used to dispose the discovery request listener.
        /// </summary>
        public static void DisposeDiscoveryListener()
        {
           if (discoveryListenerInstance != null)
           {
               discoveryListenerInstance.Dispose();
               discoveryListenerInstance = null;
           }
        }

        /// <summary>
        /// A method is used to get necessary internal http and internal https ct_netzone instances. This values is used to response to the WOPI server so that the WOPI server can use this test client as a WOPI client.
        /// </summary>
        /// <param name="currentTestClientName">A parameter represents the current test client name which run the test suite. This value is used to generated the WOPI client's app name.</param>
        /// <returns>A return value represents an array of ct_netzone type instances.</returns>
        private static ct_netzone[] GetNetZonesForWopiDiscoveryResponse(string currentTestClientName)
        {
            #region validate the parameter

            HelperBase.CheckInputParameterNullOrEmpty<string>(currentTestClientName, "currentTestClientName", "GetNetZoneForWopiDiscoveryResponse");

            #endregion 

            string fakedWOPIClientActionHostName = string.Format(@"{0}.com", Guid.NewGuid().ToString("N"));

            // Http Net zone
            ct_netzone httpNetZone = GetSingleNetZoneForWopiDiscoveryResponse(st_wopizone.internalhttp, currentTestClientName, fakedWOPIClientActionHostName);

            // Https Net zone
            ct_netzone httpsNetZone = GetSingleNetZoneForWopiDiscoveryResponse(st_wopizone.internalhttps, currentTestClientName, fakedWOPIClientActionHostName);

            return new ct_netzone[] { httpNetZone, httpsNetZone };
        }

        /// <summary>
        /// A method is used to generate a single ct_netzone type instance for current test client according to the netZoneType.
        /// </summary>
        /// <param name="netZoneType">A parameter represents the netZone type, it can only be set to "st_wopizone.internalhttp" or "st_wopizone.internalhttps"</param>
        /// <param name="currentTestClientName">A parameter represents the current test client name which run the test suite. This value is used to generated the WOPI client's app name.</param>
        /// <param name="fakedWOPIClientActionHostName">A parameter represents the host name for the action of the WOPI client support.</param>
        /// <returns>A return value represents the ct_netzone type instance.</returns>
        private static ct_netzone GetSingleNetZoneForWopiDiscoveryResponse(st_wopizone netZoneType, string currentTestClientName, string fakedWOPIClientActionHostName)
        {
            string transportValue = st_wopizone.internalhttp == netZoneType ? Uri.UriSchemeHttp : Uri.UriSchemeHttps;
            Random radomInstance = new Random((int)DateTime.UtcNow.Ticks & 0x0000FFFF);
            string appName = string.Format(
                                @"MSWOPITESTAPP {0} _for {1} WOPIServer_{2}",
                                radomInstance.Next(),
                                currentTestClientName,
                                netZoneType);

            Uri favIconUrlValue = new Uri(
                            string.Format(@"{0}://{1}/wv/resources/1033/FavIcon_Word.ico", transportValue, fakedWOPIClientActionHostName),
                            UriKind.Absolute);

            Uri urlsrcValueOfTextFile = new Uri(
                            string.Format(@"{0}://{1}/wv/wordviewerframe.aspx?&lt;ui=UI_LLCC&amp;&gt;&lt;rs=DC_LLCC&amp;&gt;&lt;showpagestats=PERFSTATS&amp;&gt;", transportValue, fakedWOPIClientActionHostName),
                            UriKind.Absolute);

            Uri urlsrcValueOfZipFile = new Uri(
                            string.Format(@"{0}://{1}/wv/zipviewerframe.aspx?&lt;ui=UI_LLCC&amp;&gt;&lt;rs=DC_LLCC&amp;&gt;&lt;showpagestats=PERFSTATS&amp;&gt;", transportValue, fakedWOPIClientActionHostName),
                            UriKind.Absolute);

            Uri urlsrcValueOfUsingprogid = new Uri(
                            string.Format(@"{0}://{1}/o/onenoteframe.aspx?edit=0&amp;&lt;ui=UI_LLCC&amp;&gt;&lt;rs=DC_LLCC&amp;&gt;&lt;showpagestats=PERFSTATS&amp;&gt;", transportValue, fakedWOPIClientActionHostName),
                            UriKind.Absolute);

            // Setting netZone's sub element's values
            ct_appname appElement = new ct_appname();
            appElement.name = appName;
            appElement.favIconUrl = favIconUrlValue.OriginalString;
            appElement.checkLicense = true;

            // Action element for txt file
            ct_wopiaction actionForTextFile = new ct_wopiaction();
            actionForTextFile.name = st_wopiactionvalues.view;
            actionForTextFile.ext = "txt";
            actionForTextFile.requires = "containers";
            actionForTextFile.@default = true;
            actionForTextFile.urlsrc = urlsrcValueOfTextFile.OriginalString;

            // Action element for txt file
            ct_wopiaction formeditactionForTextFile = new ct_wopiaction();
            formeditactionForTextFile.name = st_wopiactionvalues.formedit;
            formeditactionForTextFile.ext = "txt";
            formeditactionForTextFile.@default = true;
            formeditactionForTextFile.urlsrc = urlsrcValueOfTextFile.OriginalString;

            ct_wopiaction formViewactionForTextFile = new ct_wopiaction();
            formViewactionForTextFile.name = st_wopiactionvalues.formsubmit;
            formViewactionForTextFile.ext = "txt";
            formViewactionForTextFile.@default = true;
            formViewactionForTextFile.urlsrc = urlsrcValueOfTextFile.OriginalString;

            // Action element for zip file
            ct_wopiaction actionForZipFile = new ct_wopiaction();
            actionForZipFile.name = st_wopiactionvalues.view;
            actionForZipFile.ext = "zip";
            actionForZipFile.@default = true;
            actionForZipFile.urlsrc = urlsrcValueOfZipFile.OriginalString;

            // Action elements for one note.
            ct_wopiaction actionForOneNote = new ct_wopiaction();
            actionForOneNote.name = st_wopiactionvalues.view;
            actionForOneNote.ext = "one";
            actionForOneNote.requires = "cobalt";
            actionForOneNote.@default = true;
            actionForOneNote.urlsrc = urlsrcValueOfUsingprogid.OriginalString;

            // Action elements for one note.
            ct_wopiaction actionForOneNoteProg = new ct_wopiaction();
            actionForOneNoteProg.name = st_wopiactionvalues.view;
            actionForOneNoteProg.progid = progIdValue;
            actionForOneNoteProg.requires = "cobalt,containers";
            actionForOneNoteProg.@default = true;
            actionForOneNoteProg.urlsrc = urlsrcValueOfUsingprogid.OriginalString;

            // Add action elements into the app element.
            appElement.action = new ct_wopiaction[] { actionForTextFile, actionForOneNote, actionForZipFile, formeditactionForTextFile, formViewactionForTextFile, actionForOneNoteProg };

            // Add app element into the netzone element.
            ct_netzone netZoneInstance = new ct_netzone();
            netZoneInstance.app = new ct_appname[] { appElement };
            netZoneInstance.name = netZoneType;
            netZoneInstance.nameSpecified = true;
            return netZoneInstance;
        }
    }
}