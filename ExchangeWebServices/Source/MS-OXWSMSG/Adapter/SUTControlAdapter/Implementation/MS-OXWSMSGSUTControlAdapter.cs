namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSMSG SUT control adapter implementation. 
    /// </summary>
    public partial class MS_OXWSMSGSUTControlAdapter : ManagedAdapterBase, IMS_OXWSMSGSUTControlAdapter
    {
        #region Fields
        /// <summary>
        /// An instance of ExchangeServiceBinding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Initialize some variables overridden.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite Class.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            string userName = Common.GetConfigurationPropertyValue("Sender", this.Site);
            string password = Common.GetConfigurationPropertyValue("SenderPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }
        #endregion

        #region IMS_OXWSMSGSUTControlAdapter Operations
        /// <summary>
        /// Clean up all items in the Calendar, Inbox, Deleted Items, Drafts and Sent Items folders, which contain a specified subject.
        /// </summary>
        /// <param name="userName">The name of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="subject">Subject of the item to be removed.</param>
        /// <param name="folders">The folders to be cleaned up, which are delimited by ';'.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        public bool CleanupFolders(string userName, string password, string domain, string subject, string folders)
        {
            AdapterHelper.SwitchUser(userName, password, domain, this.exchangeServiceBinding, this.Site);

            foreach (string folder in folders.Split(';'))
            {
                if (!string.IsNullOrEmpty(folder))
                {
                    if (!this.CleanupFolder(folder, subject))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// The method searches the mailbox and returns the items that meet a specified search restriction.
        /// </summary>
        /// <param name="folder">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <returns>If the method succeeds, return an array of item; otherwise, return null.</returns>
        private ItemIdType[] GetItemIds(string folder, string value)
        {
            #region Construct FindItem request
            FindItemType findRequest = new FindItemType();

            if (string.IsNullOrEmpty(folder) || string.IsNullOrEmpty(value))
            {
                Site.Assert.Fail("Invalid argument: one or more invalid arguments passed to GetItemIds method.");
            }

            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType folderIdName;
            if (Enum.TryParse<DistinguishedFolderIdNameType>(folder, out folderIdName))
            {
                folderId.Id = folderIdName;
            }
            else
            {
                Site.Assert.Fail("The value of the first argument (foldIdNameType) of FindItem operation is invalid.");
            }

            findRequest.ParentFolderIds = new BaseFolderIdType[1];
            findRequest.ParentFolderIds[0] = folderId;

            PathToUnindexedFieldType itemClass = new PathToUnindexedFieldType();
            itemClass.FieldURI = UnindexedFieldURIType.itemSubject;
            ContainsExpressionType expressionType = new ContainsExpressionType();
            expressionType.Item = itemClass;
            expressionType.ContainmentMode = ContainmentModeType.Substring;
            expressionType.ContainmentModeSpecified = true;
            expressionType.ContainmentComparison = ContainmentComparisonType.IgnoreCaseAndNonSpacingCharacters;
            expressionType.ContainmentComparisonSpecified = true;
            expressionType.Constant = new ConstantValueType();
            expressionType.Constant.Value = value;

            RestrictionType restriction = new RestrictionType();
            restriction.Item = expressionType;

            findRequest.Restriction = restriction;
            #endregion

            #region Get the ids of all ItemId instances
            int counter = 0;
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            FindItemResponseType findResponse = new FindItemResponseType();

            while (counter < upperBound)
            {
                Thread.Sleep(waitTime);

                findResponse = this.exchangeServiceBinding.FindItem(findRequest);
                if (findResponse != null
                    && findResponse.ResponseMessages != null
                    && findResponse.ResponseMessages.Items != null
                    && findResponse.ResponseMessages.Items.Length > 0)
                {
                    ArrayOfRealItemsType items = ((FindItemResponseMessageType)findResponse.ResponseMessages.Items[0]).RootFolder.Item as ArrayOfRealItemsType;

                    if (items.Items != null && items.Items.Length > 0)
                    {
                        List<ItemIdType> itemIds = new List<ItemIdType>();
                        foreach (ItemType item in items.Items)
                        {
                            if (item.ItemId != null)
                            {
                                itemIds.Add(item.ItemId);
                            }
                        }

                        if (itemIds.Count > 0)
                        {
                            return itemIds.ToArray();
                        }
                    }
                }

                counter++;
            }

            Site.Log.Add(LogEntryKind.Debug, "When there is not any message found by FindItem operation, the retry count is {0}", counter);
            return null;
            #endregion
        }

        /// <summary>
        /// Clean up the items which contain a specified subject in the specified folder.
        /// </summary>
        /// <param name="folder">The folder to be cleaned up.</param>
        /// <param name="subject">Subject of the item to be removed.</param>
        /// <returns>If succeed, return true; otherwise, return false.</returns>
        private bool CleanupFolder(string folder, string subject)
        {
            ItemIdType[] items = this.GetItemIds(folder, subject);

            if (items != null && items.Length > 0)
            {
                DeleteItemType deleteItem = new DeleteItemType();
                deleteItem.ItemIds = items;
                deleteItem.DeleteType = DisposalType.HardDelete;
                deleteItem.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
                deleteItem.SendMeetingCancellationsSpecified = true;

                DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(deleteItem);

                if (response.ResponseMessages.Items != null)
                {
                    foreach (ResponseMessageType item in response.ResponseMessages.Items)
                    {
                        if (item.ResponseClass != ResponseClassType.Success)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }
        #endregion
    }
}