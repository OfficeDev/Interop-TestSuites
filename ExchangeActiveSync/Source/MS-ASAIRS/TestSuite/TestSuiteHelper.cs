//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        #region Build command request
        /// <summary>
        /// Build a Sync command request.
        /// </summary>
        /// <param name="syncKey">The current sync key.</param>
        /// <param name="collectionId">The collection id which to sync with.</param>
        /// <param name="commands">The sync commands.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>A Sync command request.</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, object[] commands, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            SyncRequest request = new SyncRequest
            {
                RequestData =
                {
                    Collections = new Request.SyncCollection[]
                    {
                        new Request.SyncCollection()
                        {
                            SyncKey = syncKey,
                            CollectionId = collectionId
                        }
                    }
                }
            };

            request.RequestData.Collections[0].Commands = commands;

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType1> itemsElementName = new List<Request.ItemsChoiceType1>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType1.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType1.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType1.BodyPartPreference);
                }
            }

            if (items.Count > 0)
            {
                request.RequestData.Collections[0].Options = new Request.Options[]
                {
                    new Request.Options()
                    {
                        ItemsElementName = itemsElementName.ToArray(),
                        Items = items.ToArray()
                    }
                };
            }

            return request;
        }

        /// <summary>
        /// Build an ItemOperations command request.
        /// </summary>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="serverId">The server id of the mail.</param>
        /// <param name="fileReference">The file reference of the attachment.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>An ItemOperations command request.</returns>
        internal static ItemOperationsRequest CreateItemOperationsRequest(string collectionId, string serverId, string fileReference, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            ItemOperationsRequest request = new ItemOperationsRequest { RequestData = new Request.ItemOperations() };
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch { Store = SearchName.Mailbox.ToString() };

            if (fileReference != null)
            {
                fetch.FileReference = fileReference;
            }
            else
            {
                fetch.CollectionId = collectionId;
                fetch.ServerId = serverId;
            }

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType4> itemsElementName = new List<Request.ItemsChoiceType4>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType4.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType4.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType4.BodyPartPreference);
                }
            }

            if (items.Count > 0)
            {
                fetch.Options = new Request.ItemOperationsFetchOptions()
                {
                    ItemsElementName = itemsElementName.ToArray(),
                    Items = items.ToArray()
                };
            }

            request.RequestData.Items = new object[] { fetch };

            return request;
        }

        /// <summary>
        /// Build a Search request.
        /// </summary>
        /// <param name="query">The query string.</param>
        /// <param name="collectionId">The collection id of searched folder.</param>
        /// <param name="conversationId">The conversation for which to search.</param>
        /// <param name="bodyPreferences">The bodyPreference in the options element.</param>
        /// <param name="bodyPartPreferences">The bodyPartPreference in the options element.</param>
        /// <returns>A Search command request.</returns>
        internal static SearchRequest CreateSearchRequest(string query, string collectionId, string conversationId, Request.BodyPreference[] bodyPreferences, Request.BodyPartPreference[] bodyPartPreferences)
        {
            SearchRequest request = new SearchRequest
            {
                RequestData =
                {
                    Items = new Request.SearchStore[]
                    {
                        new Request.SearchStore()
                        {
                            Name = SearchName.Mailbox.ToString(),
                            Query = new Request.queryType()
                            {
                                Items = new object[]
                                {
                                    new Request.queryType()
                                    {
                                        Items = new object[]
                                        {
                                            collectionId,
                                            query,
                                            conversationId
                                        },
                                        ItemsElementName = new Request.ItemsChoiceType5[]
                                        {
                                            Request.ItemsChoiceType5.CollectionId,
                                            Request.ItemsChoiceType5.FreeText,
                                            Request.ItemsChoiceType5.ConversationId
                                        }
                                    }
                                },
                                ItemsElementName = new Request.ItemsChoiceType5[]
                                {
                                    Request.ItemsChoiceType5.And
                                }
                            }
                        }
                    }
                }
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType6> itemsElementName = new List<Request.ItemsChoiceType6>();

            if (bodyPreferences != null)
            {
                foreach (Request.BodyPreference bodyPreference in bodyPreferences)
                {
                    items.Add(bodyPreference);
                    itemsElementName.Add(Request.ItemsChoiceType6.BodyPreference);

                    // Include the MIMESupport element in request to retrieve the MIME body
                    if (bodyPreference.Type == 4)
                    {
                        items.Add((byte)2);
                        itemsElementName.Add(Request.ItemsChoiceType6.MIMESupport);
                    }
                }
            }

            if (bodyPartPreferences != null)
            {
                foreach (Request.BodyPartPreference bodyPartPreference in bodyPartPreferences)
                {
                    items.Add(bodyPartPreference);
                    itemsElementName.Add(Request.ItemsChoiceType6.BodyPartPreference);
                }
            }

            items.Add(string.Empty);
            itemsElementName.Add(Request.ItemsChoiceType6.RebuildResults);
            items.Add("0-9");
            itemsElementName.Add(Request.ItemsChoiceType6.Range);
            items.Add(string.Empty);
            itemsElementName.Add(Request.ItemsChoiceType6.DeepTraversal);

            request.RequestData.Items[0].Options = new Request.Options1()
            {
                ItemsElementName = itemsElementName.ToArray(),
                Items = items.ToArray()
            };

            return request;
        }
        #endregion

        /// <summary>
        /// Create an instance of SyncCollection.
        /// </summary>
        /// <param name="syncKey">Specify the synchronization key obtained from the last sync command response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized, which can be returned by ActiveSync FolderSync command.</param>
        /// <returns>An instance of SyncCollection.</returns>
        internal static Request.SyncCollection CreateSyncCollection(string syncKey, string collectionId)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                SyncKey = syncKey,
                GetChanges = true,
                GetChangesSpecified = true,
                CollectionId = collectionId,
                WindowSize = "100"
            };

            return syncCollection;
        }

        #region Create MIME for SendMail command
        /// <summary>
        /// Create MIME for SendMail command.
        /// </summary>
        /// <param name="type">The email message body type.</param>
        /// <param name="from">The email address of sender.</param>
        /// <param name="to">The email address of recipient.</param>
        /// <param name="subject">The email subject.</param>
        /// <param name="body">The email body content.</param>
        /// <returns>A MIME for SendMail command.</returns>
        internal static string CreateMIME(EmailType type, string from, string to, string subject, string body)
        {
            string mime = null;
            string winmailData = null;

            // Create a plain text MIME
            if (type == EmailType.Plaintext)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/plain; charset=""us-ascii""
MIME-Version: 1.0

{3}
";
            }

            // Create an HTML MIME
            if (type == EmailType.HTML)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/html; charset=""us-ascii""
MIME-Version: 1.0

<html>
<body>
<font color=""blue"">{3}</font>
</body>
</html>
";
            }

            // Create a MIME with normal attachment
            if (type == EmailType.NormalAttachment)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/mixed; boundary=""_boundary_""; type=""text/html""
MIME-Version: 1.0

--_boundary_
Content-Type: text/html; charset=""iso-8859-1""
Content-Transfer-Encoding: quoted-printable

<html><body>{3}<img width=""128"" height=""94"" id=""Picture_x0020_1"" src=""cid:i=
mage001.jpg@01CC1FB3.2053ED80"" alt=""Description: cid:ebdc14bd-deb4-4816-b=
00b-6e2a46097d17""></body></html>

--_boundary_
Content-Type: image/jpeg; name=""number1.jpg""
Content-ID: {4}
Content-Description: number1.jpg
Content-Disposition: inline; size=4; filename=""number1.jpg""
Content-Location: <cid:ebdc14bd-deb4-4816-b00b-6e2a46097d17>
Content-Transfer-Encoding: base64

MQ==

--_boundary_--
";
            }

            // Create a MIME with embedded attachment
            if (type == EmailType.EmbeddedAttachment)
            {
                mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: multipart/mixed; boundary=""_boundary_""; type=""text/html""
MIME-Version: 1.0

--_boundary_
Content-Type: text/html; charset=""iso-8859-1""
Content-Transfer-Encoding: quoted-printable

<html><body>{3}</body></html>

--_boundary_
Content-Type: message/rfc822; name=""Embedded mail""
Content-Description: Embedded mail
Content-Disposition: attachment; size=4; filename=""Embedded mail""
Content-Transfer-Encoding: base64

MQ==

--_boundary_--
";
            }

            // Create a MIME with OLE attachment
            if (type == EmailType.AttachOLE)
            {
                winmailData = Convert.ToBase64String(File.ReadAllBytes("winmail.dat"));

                // Split lines, for the maximum length of each line in MIME is no more than 76 characters
                for (int i = 1; i < winmailData.Length / 76; i++)
                {
                    winmailData = winmailData.Insert((76 * i) - 1, "\r\n");
                }

                // The string "contoso.com" is just a sample domain name, it has no relationship to the domain configured in deployment.ptfconfig file, and any changes of this string will lead to the update of winmail.dat file.
                mime =
@"From: {0}
To: {1}
Subject: {2}
MIME-Version: 1.0
Content-Type: multipart/mixed;
    boundary=""_boundary_""
X-MS-TNEF-Correlator: <15CFAB655027B944AD65A26C7A6F2D7A0126D4994B34@DC01.contoso.com>

{3} 

--_boundary_
Content-Type: text/plain;
    charset=""us-ascii""
Content-Transfer-Encoding: 7bit

{3}  

--_boundary_
Content-Type: application/ms-tnef;
    name=""winmail.dat""
Content-Transfer-Encoding: base64
Content-Disposition: attachment;
    filename=""winmail.dat""

{5}

--_boundary_--";
            }

            return Common.FormatString(mime, from, to, subject, body, Guid.NewGuid().ToString(), winmailData);
        }
        #endregion

        /// <summary>
        /// Get the specified email item from the sync add response by using the subject.
        /// </summary>
        /// <param name="syncStore">The sync result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>Return the specified email item.</returns>
        internal static DataStructures.Sync GetSyncAddItem(DataStructures.SyncStore syncStore, string subject)
        {
            DataStructures.Sync item = null;

            if (syncStore.AddElements != null)
            {
                foreach (DataStructures.Sync syncItem in syncStore.AddElements)
                {
                    if (syncItem.Email.Subject == subject)
                    {
                        item = syncItem;
                        break;
                    }

                    if (syncItem.Contact.FileAs == subject)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Get the email item from the Search response by using the subject as the search criteria.
        /// </summary>
        /// <param name="searchStore">The Search command result.</param>
        /// <param name="subject">The email subject.</param>
        /// <returns>The email item corresponds to the specified subject.</returns>
        internal static DataStructures.Search GetSearchItem(DataStructures.SearchStore searchStore, string subject)
        {
            DataStructures.Search searchItem = null;
            if (searchStore.Results.Count > 0)
            {
                foreach (DataStructures.Search item in searchStore.Results)
                {
                    if (item.Email.Subject == subject)
                    {
                        searchItem = item;
                        break;
                    }
                }
            }

            return searchItem;
        }

        /// <summary>
        /// Truncate data according to the specified length.
        /// </summary>
        /// <param name="originalData">The original data.</param>
        /// <param name="length">The length of the byte array.</param>
        /// <returns>The truncated data.</returns>
        internal static string TruncateData(string originalData, int length)
        {
            byte[] bytes = System.Text.Encoding.ASCII.GetBytes(originalData);
            byte[] truncatedBytes = new byte[length];
            for (int i = 0; i < length; i++)
            {
                truncatedBytes[i] = bytes[i];
            }

            return System.Text.Encoding.ASCII.GetString(truncatedBytes);
        }

        /// <summary>
        /// Get the inner text of specified element.
        /// </summary>
        /// <param name="lastRawResponse">The raw xml response.</param>
        /// <param name="parentNodeName">The parent element of the specified node.</param>
        /// <param name="nodeName">The name of the node.</param>
        /// <param name="subject">The subject of the specified item.</param>
        /// <returns>The inner text of the specified element.</returns>
        internal static string GetDataInnerText(XmlElement lastRawResponse, string parentNodeName, string nodeName, string subject)
        {
            string data = null;
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(lastRawResponse.OuterXml);
            XmlNodeList subjectElementNodes = doc.SelectNodes("//*[name()='Subject']");
            for (int i = 0; i < subjectElementNodes.Count; i++)
            {
                if (subjectElementNodes[i].InnerText == subject)
                {
                    XmlNodeList bodyElementNodes = doc.SelectNodes("//*[name()='" + parentNodeName + "']");
                    XmlNodeList dataElementNodes = bodyElementNodes[i].SelectNodes("*[name()='" + nodeName + "']");
                    data = dataElementNodes[0].InnerText;
                    break;
                }
            }

            return data;
        }
    }
}