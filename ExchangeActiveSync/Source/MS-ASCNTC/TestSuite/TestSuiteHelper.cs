//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A static class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// Create MIME for SendMail command.
        /// </summary>
        /// <param name="from">The email address of sender.</param>
        /// <param name="to">The email address of recipient.</param>
        /// <param name="subject">The email subject.</param>
        /// <param name="body">The email body content.</param>
        /// <returns>A MIME for SendMail command.</returns>
        internal static string CreateMIME(string from, string to, string subject, string body)
        {
            // Create a plain text MIME
            string mime =
@"From: {0}
To: {1}
Subject: {2}
Content-Type: text/plain; charset=""us-ascii""
MIME-Version: 1.0

{3}
";
            return Common.FormatString(mime, from, to, subject, body);
        }

        /// <summary>
        /// Create an initial Sync request by using the specified collection Id.
        /// </summary>
        /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
        /// <param name="supportedElements">The elements in Supported element.</param>
        /// <returns>The SyncRequest instance.</returns>
        internal static SyncRequest CreateInitialSyncRequest(string collectionId, Request.Supported supportedElements)
        {
            Request.SyncCollection syncCollection = new Request.SyncCollection
            {
                CollectionId = collectionId,
                SyncKey = "0",
                Supported = supportedElements
            };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Create a Sync Add request.
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last Sync response</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized.</param>
        /// <param name="applicationData">The data used to specify the Add element for Sync command.</param>
        /// <returns>The SyncRequest instance.</returns>
        internal static SyncRequest CreateSyncAddRequest(string syncKey, string collectionId, Request.SyncCollectionAddApplicationData applicationData)
        {
            SyncRequest syncAddRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, null);
            Request.SyncCollectionAdd add = new Request.SyncCollectionAdd
            {
                ClientId = Guid.NewGuid().ToString("N"),
                ApplicationData = applicationData
            };

            List<object> commandList = new List<object> { add };

            syncAddRequest.RequestData.Collections[0].Commands = commandList.ToArray();

            return syncAddRequest;
        }

        /// <summary>
        /// Create a Sync Change request by using the specified sync key, folder collectionId and change application data.
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last Sync response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized.</param>
        /// <param name="changeData">The data used to specify the Change element for Sync command.</param>
        /// <returns>The SyncRequest instance.</returns>
        internal static SyncRequest CreateSyncChangeRequest(string syncKey, string collectionId, Request.SyncCollectionChange changeData)
        {
            Request.SyncCollection syncCollection = CreateSyncCollection(syncKey, collectionId);
            syncCollection.Commands = new object[] { changeData };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Create a generic Sync request without command references by using the specified sync key, folder collectionId and body preference option.
        /// </summary>
        /// <param name="syncKey">Specify the sync key obtained from the last Sync response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized.</param>
        /// <param name="bodyPreference">Sets preference information related to the type and size of information for body.</param>
        /// <returns>The SyncRequest instance.</returns>
        internal static SyncRequest CreateSyncRequest(string syncKey, string collectionId, Request.BodyPreference bodyPreference)
        {
            Request.SyncCollection syncCollection = CreateSyncCollection(syncKey, collectionId);

            Request.Options syncOptions = new Request.Options();
            List<object> syncOptionItems = new List<object>();
            List<Request.ItemsChoiceType1> syncOptionItemsName = new List<Request.ItemsChoiceType1>();

            if (null != bodyPreference)
            {
                syncOptionItemsName.Add(Request.ItemsChoiceType1.BodyPreference);
                syncOptionItems.Add(bodyPreference);

                // when body format is mime (Refer to  [MS-ASAIRS] 2.2.2.22 Type)
                if (bodyPreference.Type == 0x4)
                {
                    syncOptionItemsName.Add(Request.ItemsChoiceType1.MIMESupport);

                    // '2' indicates server sends MIME data for all messages but not S/MIME messages only.
                    syncOptionItems.Add((byte)0x2);
                }
            }

            syncOptions.Items = syncOptionItems.ToArray();
            syncOptions.ItemsElementName = syncOptionItemsName.ToArray();
            syncCollection.Options = new Request.Options[] { syncOptions };

            return Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
        }

        /// <summary>
        /// Create an instance of SyncCollection
        /// </summary>
        /// <param name="syncKey">Specify the synchronization key obtained from the last Sync command response.</param>
        /// <param name="collectionId">Specify the serverId of the folder to be synchronized.</param>
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

        /// <summary>
        /// Create a Search request.
        /// </summary>
        /// <param name="query">The query string.</param>
        /// <param name="collectionId">The collection id of searched folder.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <returns>A Search command request.</returns>
        internal static SearchRequest CreateSearchRequest(string query, string collectionId, Request.BodyPreference bodyPreference)
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
                                            query
                                        },
                                        ItemsElementName = new Request.ItemsChoiceType5[]
                                        {
                                            Request.ItemsChoiceType5.CollectionId,
                                            Request.ItemsChoiceType5.FreeText
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

            if (bodyPreference != null)
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

        /// <summary>
        /// Create an ItemOperations command request.
        /// </summary>
        /// <param name="collectionId">The collection id.</param>
        /// <param name="serverId">The serverId of the item.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <param name="schema">Sets the schema information.</param>
        /// <returns>An ItemOperations command request.</returns>
        internal static ItemOperationsRequest CreateItemOperationsRequest(string collectionId, string serverId, Request.BodyPreference bodyPreference, Request.Schema schema)
        {
            ItemOperationsRequest request = new ItemOperationsRequest { RequestData = new Request.ItemOperations() };
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                Store = SearchName.Mailbox.ToString(),
                CollectionId = collectionId,
                ServerId = serverId
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType4> itemsElementName = new List<Request.ItemsChoiceType4>();

            if (null != schema)
            {
                itemsElementName.Add(Request.ItemsChoiceType4.Schema);
                items.Add(schema);
            }

            if (null != bodyPreference)
            {
                itemsElementName.Add(Request.ItemsChoiceType4.BodyPreference);
                items.Add(bodyPreference);

                if (bodyPreference.Type == 0x4)
                {
                    itemsElementName.Add(Request.ItemsChoiceType4.MIMESupport);

                    // '2' indicates server sends MIME data for all messages but not S/MIME messages only
                    items.Add((byte)0x2);
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
        /// Get the specified email item from the Sync Add response by using the subject/FileAs.
        /// </summary>
        /// <param name="syncStore">The Sync result.</param>
        /// <param name="fileAs">The contact FileAs.</param>
        /// <returns>Return the specified email item.</returns>
        internal static DataStructures.Sync GetSyncAddItem(DataStructures.SyncStore syncStore, string fileAs)
        {
            DataStructures.Sync item = null;

            if (syncStore.AddElements.Count != 0)
            {
                foreach (DataStructures.Sync syncItem in syncStore.AddElements)
                {
                    if (syncItem.Contact.FileAs == fileAs)
                    {
                        item = syncItem;
                        break;
                    }

                    if (syncItem.Email.Subject == fileAs)
                    {
                        item = syncItem;
                        break;
                    }
                }
            }

            return item;
        }

        /// <summary>
        /// Get the specified email item from the Sync Change response by using the subject.
        /// </summary>
        /// <param name="syncStore">The Sync result.</param>
        /// <param name="fileAs">The contact FileAs.</param>
        /// <returns>Return the specified email item.</returns>
        internal static DataStructures.Sync GetSyncChangeItem(DataStructures.SyncStore syncStore, string fileAs)
        {
            DataStructures.Sync item = null;

            if (syncStore.ChangeElements.Count != 0)
            {
                foreach (DataStructures.Sync syncItem in syncStore.ChangeElements)
                {
                    if (syncItem.Contact.FileAs == fileAs)
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
        /// <param name="fileAs">The FileAs of the contact.</param>
        /// <returns>The email item corresponds to the specified subject.</returns>
        internal static DataStructures.Search GetSearchItem(DataStructures.SearchStore searchStore, string fileAs)
        {
            DataStructures.Search searchItem = null;
            if (searchStore.Results.Count > 0)
            {
                foreach (DataStructures.Search item in searchStore.Results)
                {
                    if (item.Contact.FileAs == fileAs)
                    {
                        searchItem = item;
                        break;
                    }
                }
            }

            return searchItem;
        }

        /// <summary>
        /// Check if the response message only contains the specified element in the specified xml tag.
        /// </summary>
        /// <param name="rawResponseXml">The raw xml of the response returned by SUT</param>
        /// <param name="tagName">The name of the specified xml tag.</param>
        /// <param name="elementName">The element name that the raw xml should contain.</param>
        /// <returns>If the response only contains the specified element, return true; otherwise, false.</returns>
        internal static bool IsOnlySpecifiedElementExist(XmlElement rawResponseXml, string tagName, string elementName)
        {
            bool isOnlySpecifiedElement = false;
            if (rawResponseXml != null)
            {
                XmlNodeList nodes = rawResponseXml.GetElementsByTagName(tagName);
                foreach (XmlNode node in nodes)
                {
                    if (node.HasChildNodes)
                    {
                        XmlNodeList children = node.ChildNodes;
                        if (children.Count > 0)
                        {
                            foreach (XmlNode child in children)
                            {
                                if (string.Equals(child.Name, elementName))
                                {
                                    isOnlySpecifiedElement = true;
                                }
                                else
                                {
                                    isOnlySpecifiedElement = false;
                                    break;
                                }
                            }
                        }
                        else
                        {
                            isOnlySpecifiedElement = false;
                        }
                    }
                    else
                    {
                        isOnlySpecifiedElement = false;
                    }
                }
            }

            return isOnlySpecifiedElement;
        }
    }
}