namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the method defined in interface IMS_OXWSSRCHSUTControlAdapter. 
    /// </summary>
    public partial class MS_OXWSSRCHSUTControlAdapter : ManagedAdapterBase, IMS_OXWSSRCHSUTControlAdapter
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

        #region IMS_OXWSSRCHSUTControlAdapter Operations
        /// <summary>
        /// The operation searches the mailbox and returns the result whether one or more valid items are found.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the operation succeeds, return true; otherwise, return false.</returns>
        public bool FindItem(string userName, string password, string domain, string folderName, string value, string field)
        {
            AdapterHelper.SwitchUser(userName, password, domain, this.exchangeServiceBinding, this.Site);

            // Construct a request for FindItem operation.
            FindItemType findRequest = this.ConstructFindItemRequest(folderName, value, field);

            #region Invoke FindItem operation
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
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        if (item.ResponseClass == ResponseClassType.Success)
                        {
                            FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                            if (findItem.RootFolder.Item != null)
                            {
                                ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                                if (realItems.Items != null && realItems.Items.Length > 0)
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }

                counter++;
            }

            Site.Log.Add(LogEntryKind.Debug, "Even after retrying {0} times, operation FindItem could not find any message.", counter);
            return false;
            #endregion
        }

        /// <summary>
        /// The operation searches the specified user's mailbox and returns the result whether the valid items are found after the MoveItem or DeleteItem operation completed.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="domain">Domain of the user.</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the item existed in the specific folder, return true; otherwise, return false.</returns>
        public bool IsItemAvailableAfterMoveOrDelete(string userName, string password, string domain, string folderName, string value, string field)
        {
            AdapterHelper.SwitchUser(userName, password, domain, this.exchangeServiceBinding, this.Site);
            
            // Construct a request for FindItem operation.
            FindItemType findRequest = this.ConstructFindItemRequest(folderName, value, field);

            #region Invoke FindItem operation
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
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        if (item.ResponseClass == ResponseClassType.Success)
                        {
                            FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                            if (findItem.RootFolder.Item != null)
                            {
                                ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                                if (realItems.Items != null && realItems.Items.Length > 0)
                                {
                                    counter++;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
            }

            Site.Log.Add(LogEntryKind.Debug, "Even after retrying {0} times, Message with specified subject is still available, this means previous call to MoveItem or DeleteItem has not worked yet.", counter);
            return true;
            #endregion
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// The operation construct a request for FindItem operation.
        /// </summary>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>The request of FindItem operation which constructed with special folder name, search restriction and referenced field URI</returns>
        protected FindItemType ConstructFindItemRequest(string folderName, string value, string field)
        {
            FindItemType findRequest = new FindItemType();
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType folderIdName;
            if (Enum.TryParse<DistinguishedFolderIdNameType>(folderName, out folderIdName))
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
            UnindexedFieldURIType fieldURI;
            if (Enum.TryParse<UnindexedFieldURIType>(field, out fieldURI))
            {
                // set search field.
                itemClass.FieldURI = fieldURI;
            }
            else
            {
                Site.Assert.Fail("The value of the second argument (fieldURIType) of FindItem operation is invalid.");
            }

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

            if (!string.IsNullOrEmpty(value))
            {
                findRequest.Restriction = restriction;
            }

            return findRequest;
        }
        #endregion
    }
}