//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASHTTP.
    /// </summary>
    public partial class MS_ASHTTPAdapter : ManagedAdapterBase, IMS_ASHTTPAdapter
    {
        #region Fields
        /// <summary>
        /// ActiveSyncClient instance.
        /// </summary>
        private ActiveSyncClient activeSyncClient;
        #endregion

        #region IMS_ASHTTPAdapter Properties
        /// <summary>
        /// Gets the XML request sent to protocol SUT
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the XML response received from protocol SUT
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASHTTP";

            // Merge the configuration.
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }

        #region MS-ASHTTP protocol methods
        /// <summary>
        /// Send HTTP POST request to the server and get the response.
        /// </summary>
        /// <param name="commandName">The name of the command to send.</param>
        /// <param name="commandParameters">The command parameters.</param>
        /// <param name="requestBody">The plain text request.</param>
        /// <returns>The plain text response.</returns>
        public SendStringResponse HTTPPOST(CommandName commandName, IDictionary<CmdParameterName, object> commandParameters, string requestBody)
        {
            SendStringResponse postResponse = this.activeSyncClient.SendStringRequest(commandName, commandParameters, requestBody);
            Site.Assert.IsNotNull(postResponse, "The HTTP POST response returned from server should not be null.");
            this.VerifyHTTPPOSTResponse(postResponse);
            this.VerifyTransportType();
            return postResponse;
        }

        /// <summary>
        /// Send HTTP OPTIONS request to the server and get the response.
        /// </summary>
        /// <returns>The HTTP OPTIONS response.</returns>
        public OptionsResponse HTTPOPTIONS()
        {
            OptionsResponse optionsResponse = this.activeSyncClient.Options();
            Site.Assert.IsNotNull(optionsResponse, "The HTTP OPTIONS response returned from server should not be null.");
            this.VerifyHTTPOPTIONSResponse(optionsResponse);
            this.VerifyTransportType();
            return optionsResponse;
        }
        #endregion

        #region Update activeSyncClient properties.
        /// <summary>
        /// Configure the fields in request line or request headers besides command name and command parameters.
        /// </summary>
        /// <param name="requestPrefixFields">The fields in request line or request headers which need to be configured besides command name and command parameters.</param>
        public void ConfigureRequestPrefixFields(IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields)
        {
            if (requestPrefixFields != null)
            {
                foreach (KeyValuePair<HTTPPOSTRequestPrefixField, string> requestPrefixField in requestPrefixFields)
                {
                    if (requestPrefixField.Key == HTTPPOSTRequestPrefixField.QueryValueType)
                    {
                        this.activeSyncClient.QueryValueType = (QueryValueType)Enum.Parse(typeof(QueryValueType), requestPrefixField.Value, true);
                    }
                    else
                    {
                        this.activeSyncClient.GetType().GetProperty(requestPrefixField.Key.ToString()).SetValue(
                            this.activeSyncClient, 
                            Convert.ChangeType(requestPrefixField.Value, this.activeSyncClient.GetType().GetProperty(requestPrefixField.Key.ToString()).PropertyType),
                            null);
                    }
                }
            }
        }
        #endregion
    }
}