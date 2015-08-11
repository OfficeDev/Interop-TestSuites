namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the methods defined in interface IMS_OXWSFOLDSUTControlAdapter.
    /// </summary>
    public class MS_OXWSFOLDSUTControlAdapter : ManagedAdapterBase, IMS_OXWSFOLDSUTControlAdapter
    {
        #region Fields
        /// <summary>
        /// The endpoint url of Exchange Web Service.
        /// </summary>
        private string url;

        /// <summary>
        /// The password for userName used to access web service.
        /// </summary>
        private string password;

        /// <summary>
        /// The user name used to access web service.
        /// </summary>
        private string userName;

        /// <summary>
        /// The domain of server.
        /// </summary>
        private string domain;

        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSSYNC";

            // Merge configuration files.
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            this.password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSFOLDSUTControlAdapter Operations
        /// <summary>
        /// Log on to a mailbox with a specified user account and create two different-level subfolders in the specified parent folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="firstLevelSubFolderName">Name of the first level sub folder which will be created under the parent folder.</param>
        /// <param name="secondLevelSubFolderName">Name of the second level sub folder which will be created under the first level sub folder.</param>
        /// <returns>If the two level sub folders are created successfully, return true; otherwise, return false.</returns>
        public bool CreateSubFolders(string userName, string userPassword, string userDomain, string parentFolderName, string firstLevelSubFolderName, string secondLevelSubFolderName)
        {
            // Log on mailbox with specified user account(userName, userPassword, userDomain).
            bool isLoged = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isLoged,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Initialize variables
            FolderIdType folderId = null;
            CreateFolderType createFolderRequest = new CreateFolderType();
            string folderClassName = null;
            DistinguishedFolderIdNameType parentFolderIdName = (DistinguishedFolderIdNameType)Enum.Parse(typeof(DistinguishedFolderIdNameType), parentFolderName, true);

            // Define different folder class name according to different parent folder.
            switch (parentFolderIdName)
            {
                case DistinguishedFolderIdNameType.inbox:
                    folderClassName = "IPF.Note";
                    break;
                case DistinguishedFolderIdNameType.contacts:
                    folderClassName = "IPF.Contact";
                    break;
                case DistinguishedFolderIdNameType.calendar:
                    folderClassName = "IPF.Appointment";
                    break;
                case DistinguishedFolderIdNameType.tasks:
                    folderClassName = "IPF.Task";
                    break;
                default:
                    Site.Assume.Fail(string.Format("The parent folder name '{0}' is invalid. Valid values are: inbox, contacts, calendar or tasks.", parentFolderName));
                    break;
            }

            // Set parent folder ID.
            createFolderRequest.ParentFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType parentFolder = new DistinguishedFolderIdType();
            parentFolder.Id = parentFolderIdName;
            createFolderRequest.ParentFolderId.Item = parentFolder;

            // Set Display Name and Folder Class for the folder to be created.
            FolderType folderProperties = new FolderType();
            folderProperties.DisplayName = firstLevelSubFolderName;
            folderProperties.FolderClass = folderClassName;

            createFolderRequest.Folders = new BaseFolderType[1];
            createFolderRequest.Folders[0] = folderProperties;

            bool isSubFolderCreated = false;

            // Invoke CreateFolder operation and get the response.
            CreateFolderResponseType createFolderResponse = this.exchangeServiceBinding.CreateFolder(createFolderRequest);

            if (createFolderResponse != null && ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass)
            {
                // If the first level sub folder is created successfully, save the folder ID of it.
                folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
                FolderType created = new FolderType() { DisplayName = folderProperties.DisplayName, FolderClass = folderClassName, FolderId = folderId };
                AdapterHelper.CreatedFolders.Add(created);
            }

            // Create another sub folder under the created folder above.
            if (folderId != null)
            {
                createFolderRequest.ParentFolderId.Item = folderId;
                folderProperties.DisplayName = secondLevelSubFolderName;

                createFolderResponse = this.exchangeServiceBinding.CreateFolder(createFolderRequest);

                if (createFolderResponse != null && ResponseClassType.Success == createFolderResponse.ResponseMessages.Items[0].ResponseClass)
                {
                    // If the two level sub folders are created successfully, return true; otherwise, return false.
                    isSubFolderCreated = true;
                    folderId = ((FolderInfoResponseMessageType)createFolderResponse.ResponseMessages.Items[0]).Folders[0].FolderId;
                    FolderType created = new FolderType() { DisplayName = folderProperties.DisplayName, FolderClass = folderClassName, FolderId = folderId };
                    AdapterHelper.CreatedFolders.Add(created);
                }
            }

            return isSubFolderCreated;
        }

        /// <summary>
        /// Log on to a mailbox with a specified user account and create a search folder.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="searchFolderName">Name of the search folder.</param>
        /// <param name="searchText">Search text of the search folder.</param>
        /// <returns>If the search folder is created successfully, return true; otherwise, return false.</returns>
        public bool CreateSearchFolder(string userName, string userPassword, string userDomain, string searchFolderName, string searchText)
        {
            // Log on mailbox with specified user account(userName, userPassword, userDomain).
            bool isLoged = AdapterHelper.SwitchUser(userName, userPassword, userDomain, this.exchangeServiceBinding, this.Site);
            Site.Assert.IsTrue(
                isLoged,
                string.Format("Log on mailbox with the UserName: {0}, Password: {1}, Domain: {2} should be successful.", userName, userPassword, userDomain));

            // Create the request.
            CreateFolderType createFolder = new CreateFolderType();
            SearchFolderType[] folderArray = new SearchFolderType[1];
            SearchFolderType searchFolder = new SearchFolderType();

            // Use the following search filter to get all mail in the Inbox with the word searchText in the subject line.
            searchFolder.SearchParameters = new SearchParametersType();
            searchFolder.SearchParameters.Traversal = SearchFolderTraversalType.Deep;
            searchFolder.SearchParameters.TraversalSpecified = true;
            searchFolder.SearchParameters.BaseFolderIds = new DistinguishedFolderIdType[4];

            // Create a distinguished folder Identified of the inbox folder.
            DistinguishedFolderIdType inboxFolder = new DistinguishedFolderIdType();
            inboxFolder.Id = new DistinguishedFolderIdNameType();
            inboxFolder.Id = DistinguishedFolderIdNameType.inbox;
            searchFolder.SearchParameters.BaseFolderIds[0] = inboxFolder;
            DistinguishedFolderIdType contactType = new DistinguishedFolderIdType();
            contactType.Id = new DistinguishedFolderIdNameType();
            contactType.Id = DistinguishedFolderIdNameType.contacts;
            searchFolder.SearchParameters.BaseFolderIds[1] = contactType;
            DistinguishedFolderIdType calendarType = new DistinguishedFolderIdType();
            calendarType.Id = new DistinguishedFolderIdNameType();
            calendarType.Id = DistinguishedFolderIdNameType.calendar;
            searchFolder.SearchParameters.BaseFolderIds[2] = calendarType;
            DistinguishedFolderIdType taskType = new DistinguishedFolderIdType();
            taskType.Id = new DistinguishedFolderIdNameType();
            taskType.Id = DistinguishedFolderIdNameType.calendar;
            searchFolder.SearchParameters.BaseFolderIds[3] = taskType;

            // Use the following search filter.
            searchFolder.SearchParameters.Restriction = new RestrictionType();
            PathToUnindexedFieldType path = new PathToUnindexedFieldType();
            path.FieldURI = UnindexedFieldURIType.itemSubject;
            RestrictionType restriction = new RestrictionType();
            FieldURIOrConstantType fieldURIORConstant = new FieldURIOrConstantType();
            fieldURIORConstant.Item = new ConstantValueType();
            (fieldURIORConstant.Item as ConstantValueType).Value = searchText;
            ExistsType isEqual = new ExistsType();
            isEqual.Item = path;
            restriction.Item = isEqual;
            searchFolder.SearchParameters.Restriction = restriction;

            // Give the search folder a unique name.
            searchFolder.DisplayName = searchFolderName;
            folderArray[0] = searchFolder;

            // Create the search folder under the default Search Folder.
            TargetFolderIdType targetFolder = new TargetFolderIdType();
            DistinguishedFolderIdType searchFolders = new DistinguishedFolderIdType();
            searchFolders.Id = DistinguishedFolderIdNameType.searchfolders;
            targetFolder.Item = searchFolders;
            createFolder.ParentFolderId = targetFolder;
            createFolder.Folders = folderArray;
            bool isSearchFolderCreated = false;

            // Invoke CreateFolder operation and get the response.
            CreateFolderResponseType response = this.exchangeServiceBinding.CreateFolder(createFolder);
            if (response != null && ResponseClassType.Success == response.ResponseMessages.Items[0].ResponseClass)
            {
                // If the search folder is created successfully, return true; otherwise, return false.
                isSearchFolderCreated = true;

                searchFolder.FolderId = ((FolderInfoResponseMessageType)response.ResponseMessages.Items[0]).Folders[0].FolderId;
                AdapterHelper.CreatedFolders.Add(searchFolder);
            }

            return isSearchFolderCreated;
        }
        #endregion
    }
}