namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    internal static class TestSuiteHelper
    {
        /// <summary>
        /// Get the request of Sync command.
        /// </summary>
        /// <param name="collectionId">The collection id of the folder to sync.</param>
        /// <param name="syncKey">The SyncKey of the latest sync.</param>
        /// <param name="bodyPartPreference">The bodyPartPreference in the options element.</param>
        /// <param name="bodyPreference">The bodyPreference in the options element.</param>
        /// <param name="conversationMode">The value of ConversationMode element.</param>
        /// <returns>The request of Sync command.</returns>
        internal static SyncRequest GetSyncRequest(string collectionId, string syncKey, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference, bool conversationMode)
        {
            // Create the Sync command request.
            Request.SyncCollection[] synCollections = new Request.SyncCollection[1];
            synCollections[0] = new Request.SyncCollection { SyncKey = syncKey, CollectionId = collectionId };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType1> itemsElementName = new List<Request.ItemsChoiceType1>();

            if (bodyPartPreference != null)
            {
                items.Add(bodyPartPreference);
                itemsElementName.Add(Request.ItemsChoiceType1.BodyPartPreference);
            }

            if (bodyPreference != null)
            {
                items.Add(bodyPreference);
                itemsElementName.Add(Request.ItemsChoiceType1.BodyPreference);
            }

            if (conversationMode)
            {
                synCollections[0].ConversationMode = true;
                synCollections[0].ConversationModeSpecified = true;
                synCollections[0].Options = new Request.Options[1];
                synCollections[0].Options[0] = new Request.Options
                {
                    ItemsElementName = new Request.ItemsChoiceType1[] { Request.ItemsChoiceType1.FilterType },
                    Items = new object[] { (byte)1 }
                };
            }

            if (items.Count > 0)
            {
                synCollections[0].Options = new Request.Options[]
                {
                    new Request.Options()
                    {
                        ItemsElementName = itemsElementName.ToArray(),
                        Items = items.ToArray()
                    }
                };
            }

            return Common.CreateSyncRequest(synCollections);
        }

        /// <summary>
        /// Get the request of ItemOperations command.
        /// </summary>
        /// <param name="collectionId">The collection id of the folder to fetch.</param>
        /// <param name="serverId">The server id of the mail.</param>
        /// <param name="bodyPartPreference">The BodyPartPreference in the Options element.</param>
        /// <param name="bodyPreference">The BodyPreference in the Options element.</param>
        /// <returns>An ItemOperations command request.</returns>
        internal static ItemOperationsRequest GetItemOperationsRequest(string collectionId, string serverId, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
        {
            ItemOperationsRequest request = new ItemOperationsRequest { RequestData = new Request.ItemOperations() };
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                Store = SearchName.Mailbox.ToString(),
                CollectionId = collectionId,
                ServerId = serverId
            };

            List<object> items = new List<object>();
            List<Request.ItemsChoiceType5> itemsElementName = new List<Request.ItemsChoiceType5>();

            if (bodyPreference != null)
            {
                items.Add(bodyPreference);
                itemsElementName.Add(Request.ItemsChoiceType5.BodyPreference);
            }

            if (bodyPartPreference != null)
            {
                items.Add(bodyPartPreference);
                itemsElementName.Add(Request.ItemsChoiceType5.BodyPartPreference);
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
        /// Get the request of Search command.
        /// </summary>
        /// <param name="conversationId">The conversation for which to search.</param>
        /// <param name="bodyPartPreference">The BodyPartPreference in the Options element.</param>
        /// <param name="bodyPreference">The BodyPreference in the Options element.</param>
        /// <returns>A Search command request.</returns>
        internal static SearchRequest GetSearchRequest(string conversationId, Request.BodyPartPreference bodyPartPreference, Request.BodyPreference bodyPreference)
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
                                            "Email",
                                            conversationId
                                        },
                                        ItemsElementName = new Request.ItemsChoiceType2[]
                                        {
                                            Request.ItemsChoiceType2.Class,
                                            Request.ItemsChoiceType2.ConversationId
                                        }
                                    }
                                },
                                ItemsElementName = new Request.ItemsChoiceType2[]
                                {
                                    Request.ItemsChoiceType2.And
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
            }

            if (bodyPartPreference != null)
            {
                items.Add(bodyPartPreference);
                itemsElementName.Add(Request.ItemsChoiceType6.BodyPartPreference);
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