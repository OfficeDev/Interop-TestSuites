//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// Builds a generic Sync request without command references by using the specified sync key, folder collection ID and body preference option.
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last sync response</param>
        /// <param name="collectionId">Specify the server ID of the folder to be synchronized, which can be returned by ActiveSync FolderSync command</param>
        /// <param name="rightsManagementSupport">A boolean value specifies whether the server will decompress and decrypt rights-managed email messages before sending them to the client or not</param>
        /// <returns>Returns the Sync request instance</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, bool? rightsManagementSupport)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                CollectionId = collectionId
            };

            if (syncKey != "0")
            {
                syncCollection.GetChanges = true;
                syncCollection.GetChangesSpecified = true;
            }

            syncCollection.WindowSize = "100";

            Request.BodyPreference bodyPreference = new Request.BodyPreference { Type = 1 };

            Request.Options syncOptions = new Request.Options();
            List<object> syncOptionItems = new List<object> { bodyPreference };

            List<Request.ItemsChoiceType1> syncOptionItemsName = new List<Request.ItemsChoiceType1>
            {
                Request.ItemsChoiceType1.BodyPreference
            };

            if (rightsManagementSupport != null)
            {
                syncOptionItems.Add(rightsManagementSupport);
                syncOptionItemsName.Add(Request.ItemsChoiceType1.RightsManagementSupport);
            }

            syncOptions.Items = syncOptionItems.ToArray();
            syncOptions.ItemsElementName = syncOptionItemsName.ToArray();
            syncCollection.Options = new Request.Options[] { syncOptions };
            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Builds a SmartForward request by using the specified source folder Id, source server Id and forward mime information.
        /// </summary>
        /// <param name="sourceFolderId">Specify the folder id of original mail item being forwarded</param>
        /// <param name="sourceServerId">Specify the server Id of original mail item being forwarded</param>
        /// <param name="forwardMime">The total forward mime</param>
        /// <returns>Returns the SmartForward request instance</returns>
        internal static SmartForwardRequest CreateSmartForwardRequest(string sourceFolderId, string sourceServerId, string forwardMime)
        {
            SmartForwardRequest request = new SmartForwardRequest
            {
                RequestData = new Request.SmartForward
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source { FolderId = sourceFolderId, ItemId = sourceServerId },
                    Mime = forwardMime
                }
            };

            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, sourceFolderId
                },
                {
                    CmdParameterName.ItemId, sourceServerId
                }
            });
            return request;
        }

        /// <summary>
        /// Builds a SmartReply request by using the specified source folder Id, source server Id and reply mime information.
        /// </summary>
        /// <param name="sourceFolderId">Specify the folder id of original mail item being replied</param>
        /// <param name="sourceServerId">Specify the server Id of original mail item being replied</param>
        /// <param name="replyMime">The total reply mime</param>
        /// <returns>Returns the SmartReply request instance</returns>
        internal static SmartReplyRequest CreateSmartReplyRequest(string sourceFolderId, string sourceServerId, string replyMime)
        {
            SmartReplyRequest request = new SmartReplyRequest
            {
                RequestData = new Request.SmartReply
                {
                    ClientId = System.Guid.NewGuid().ToString(),
                    Source = new Request.Source { FolderId = sourceFolderId, ItemId = sourceServerId },
                    Mime = replyMime
                }
            };

            request.SetCommandParameters(new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.CollectionId, sourceFolderId
                },
                {
                    CmdParameterName.ItemId, sourceServerId
                }
            });
            return request;
        }

        /// <summary>
        /// Create a SendMail request.
        /// </summary>
        /// <param name="from">The sender address of mail</param>
        /// <param name="to">The recipient address of the mail</param>
        /// <param name="cc">The carbon copy address of the mail</param>
        /// <param name="bcc">The blind cc address of the mail</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="content">The content of the mail</param>
        /// <param name="templateID">A string that identifies a particular rights policy template to be applied to the outgoing message.</param>
        /// <returns>The SendMail request.</returns>
        internal static SendMailRequest CreateSendMailRequest(string from, string to, string cc, string bcc, string subject, string content, string templateID)
        {
            string clientId = System.Guid.NewGuid().ToString();
            string mime = Common.CreatePlainTextMime(from, to, cc, bcc, subject, content);
            SendMailRequest sendMailRequest = Common.CreateSendMailRequest(null, clientId, mime);
            sendMailRequest.RequestData.TemplateID = templateID;
            return sendMailRequest;
        }

        /// <summary>
        /// Creates an instance of SyncCollection
        /// </summary>
        /// <param name="syncKey">Specify the synchronization key obtained from the last sync command response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized, which can be returned by ActiveSync FolderSync command</param>
        /// <returns>An instance of SyncCollection</returns>
        internal static Request.SyncCollection CreateSyncCollection(string syncKey, string collectionId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                CollectionId = collectionId,
                DeletesAsMoves = false,
                DeletesAsMovesSpecified = true,
                WindowSize = "100"
            };

            return syncCollection;
        }

        /// <summary>
        /// Get the inner text of specified element.
        /// </summary>
        /// <param name="lastRawResponse">The raw xml response.</param>
        /// <param name="parentNodeName">The parent element of the specified node.</param>
        /// <param name="nodeName">The name of the node.</param>
        /// <param name="subject">The subject of the specified item.</param>
        /// <returns>The inner text of the specified element.</returns>
        internal static string GetElementInnerText(XmlElement lastRawResponse, string parentNodeName, string nodeName, string subject)
        {
            string data = null;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(lastRawResponse.OuterXml);
            XmlNodeList subjectElementNodes = doc.SelectNodes("//*[name()='Subject']");
            for (int i = 0; i < subjectElementNodes.Count; i++)
            {
                if (subjectElementNodes[i].InnerText == subject)
                {
                    XmlNodeList bodyElementNodes = subjectElementNodes[i].ParentNode.SelectNodes("*[name()='" + parentNodeName + "']");
                    XmlNodeList dataElementNodes = bodyElementNodes[0].SelectNodes("*[name()='" + nodeName + "']");
                    data = dataElementNodes[0].InnerText;
                    break;
                }
            }

            return data;
        }
    }
}
