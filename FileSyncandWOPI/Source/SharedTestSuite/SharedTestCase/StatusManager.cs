namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to make sure clean up the test environment.
    /// </summary>
    public class StatusManager
    {
        /// <summary>
        /// A dictionary which maintains the mapping between the file URL and the release lock function.
        /// </summary>
        private Dictionary<SharedLockKey, Action> releaseSharedLockFunctions;

        /// <summary>
        /// A dictionary which maintains the mapping between the file URL and the release lock function.
        /// </summary>
        private Dictionary<string, Action> releaseExclusiveLockFunctions;

        /// <summary>
        /// A dictionary which maintains the mapping between the file URL and the server status roll back functions.
        /// </summary>
        private Dictionary<ServerStatus, Action> documentLibraryStatusRollbackFunctions;

        /// <summary>
        /// A dictionary which maintains the mapping between the file URL and the check out functions.
        /// </summary>
        private Dictionary<string, Action> fileCheckOutRollbackFunctions;

        /// <summary>
        /// A dictionary which maintains the mapping between the file URL and removing file functions.
        /// </summary>
        private Dictionary<string, Action> removeFileFunctions;

        /// <summary>
        /// A recording to record all the error message.
        /// </summary>
        private List<string> errorMessage;

        /// <summary>
        /// A value indicates whether there is an error when do the environment clean up. 
        /// </summary>
        private bool isEnvironmentRollbackSuccess;

        /// <summary>
        /// A ITestSite instance.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// Specify the initialize context call back function.
        /// </summary>
        private Action<string, string, string, string> intializeContext;

        /// <summary>
        /// Initializes a new instance of the StatusManager class.
        /// </summary>
        /// <param name="site">Specify the ITestSite instance.</param>
        /// <param name="intializeContext">Specify the initialize context call back function.</param>
        public StatusManager(ITestSite site, Action<string, string, string, string> intializeContext)
        {
            this.releaseSharedLockFunctions = new Dictionary<SharedLockKey, Action>();
            this.releaseExclusiveLockFunctions = new Dictionary<string, Action>();
            this.fileCheckOutRollbackFunctions = new Dictionary<string, Action>();
            this.documentLibraryStatusRollbackFunctions = new Dictionary<ServerStatus, Action>();
            this.removeFileFunctions = new Dictionary<string, Action>();
            this.errorMessage = new List<string>();
            this.site = site;
            this.isEnvironmentRollbackSuccess = true;
            this.intializeContext = intializeContext;
        }

        /// <summary>
        /// The Enum of the server status.
        /// </summary>
        public enum ServerStatus
        {
            /// <summary>
            /// File needed to be checked out when getting various lock.
            /// </summary>
            CheckOutRequired,

            /// <summary>
            /// disable the coauthoring feature.
            /// </summary>
            DisableCoauth,

            /// <summary>
            /// Turn off the cell storage web service.
            /// </summary>
            DisableCellStorageWebService,

            /// <summary>
            /// Change the authentication to window based.
            /// </summary>
            DisableClaimsBasedAuthentication,

            /// <summary>
            /// Disable the versioning on the document library.
            /// </summary>
            DisableVersioning
        }

        /// <summary>
        /// The key status for update the specified records.
        /// </summary>
        public enum KeyStatus
        {
            /// <summary>
            /// Update the check out records.
            /// </summary>
            CheckOut,

            /// <summary>
            /// Update the exclusive lock records.
            /// </summary>
            ExclusiveLock,

            /// <summary>
            /// Update the uploading text file records.
            /// </summary>
            UploadTextFile
        }

        /// <summary>
        /// Roll back all the file and document library status to do the environment clean up.
        /// </summary>
        /// <returns>Return true indicate the status rollback is successful, otherwise false.</returns>
        public bool RollbackStatus()
        {
            foreach (KeyValuePair<ServerStatus, Action> pairs in this.documentLibraryStatusRollbackFunctions)
            {
                if (pairs.Key != ServerStatus.DisableClaimsBasedAuthentication)
                {
                    pairs.Value();
                }
            }

            foreach (Action rollbackFunction in this.fileCheckOutRollbackFunctions.Values)
            {
                rollbackFunction();
            }

            foreach (Action rollbackFunction in this.releaseExclusiveLockFunctions.Values)
            {
                rollbackFunction();
            }

            foreach (Action rollbackFunction in this.releaseSharedLockFunctions.Values)
            {
                rollbackFunction();
            }

            // After all the lock or check out status has been released, then can change the authentication mode to claims based authentication.
            foreach (KeyValuePair<ServerStatus, Action> pairs in this.documentLibraryStatusRollbackFunctions)
            {
                if (pairs.Key == ServerStatus.DisableClaimsBasedAuthentication)
                {
                    pairs.Value();
                }
            }

            this.releaseExclusiveLockFunctions.Clear();
            this.documentLibraryStatusRollbackFunctions.Clear();
            this.releaseSharedLockFunctions.Clear();

            return this.isEnvironmentRollbackSuccess;
        }

        /// <summary>
        /// This method is used to remove all the upload files by the records.
        /// </summary>
        /// <returns>Return true indicate the removing operation is successful, otherwise false.</returns>
        public bool CleanUpFiles()
        {
            foreach (Action function in this.removeFileFunctions.Values)
            {
                function();
            }

            this.removeFileFunctions.Clear();
            return this.isEnvironmentRollbackSuccess;
        }

        /// <summary>
        /// This method is used to generate error message report when cleaning the environment.
        /// </summary>
        /// <returns>Return the message report.</returns>
        public string GenerateErrorMessageReport()
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (string message in this.errorMessage)
            {
                sb.Append(message);
                sb.Append(Environment.NewLine);
            }

            this.errorMessage.Clear();
            return sb.ToString();
        }

        /// <summary>
        /// This method is used to record the editors table status with specified client ID and editorsTableType for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the file we use</param>
        /// <param name="clientId">Specify the client ID of the editors table.</param>
        public void RecordEditorTable(string fileUrl, string clientId)
        {
            string userName = Common.GetConfigurationPropertyValue("UserName1", this.site);
            string password = Common.GetConfigurationPropertyValue("Password1", this.site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.site);

            this.RecordEditorTable(fileUrl, clientId, userName, password, domain);
        }

        /// <summary>
        /// This method is used to record the Editors Table status with specified client ID and editorsTableType for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the file we use.</param>
        /// <param name="clientId">Specify the client ID of the editors table.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        public void RecordEditorTable(string fileUrl, string clientId, string userName, string password, string domain)
        {
            Action function = () =>
            {
                EditorsTableSubRequestType ets = new EditorsTableSubRequestType();
                ets.Type = SubRequestAttributeType.EditorsTable;
                ets.SubRequestData = new EditorsTableSubRequestDataType();
                ets.SubRequestData.ClientID = clientId;
                ets.SubRequestData.EditorsTableRequestTypeSpecified = true;
                ets.SubRequestData.EditorsTableRequestType = EditorsTableRequestTypes.LeaveEditingSession;
                ets.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString(); 
                
                IMS_FSSHTTP_FSSHTTPBAdapter adapter = site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
                this.intializeContext(fileUrl, userName, password, domain);
                EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(adapter.CellStorageRequest(fileUrl, new SubRequestType[] { ets }), 0, 0, site);

                if (!string.Equals("Success", subResponse.ErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    this.errorMessage.Add(string.Format("Failed to release the editor tables join status for the client id {1} on the file {0} using the following user: {2}/{3}and password:{4}", fileUrl, clientId, userName, domain, password));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(new SharedLockKey(fileUrl, clientId, string.Empty), function);
        }

        /// <summary>
        /// This method is used to record the schema lock status with specified client ID and schema lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the schema lock.</param>
        /// <param name="clientId">Specify the client ID of the schema lock.</param>
        /// <param name="schemaLockId">Specify the schema ID of the schema lock.</param>
        public void RecordSchemaLock(string fileUrl, string clientId, string schemaLockId)
        {
            string userName = Common.GetConfigurationPropertyValue("UserName1", this.site);
            string password = Common.GetConfigurationPropertyValue("Password1", this.site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.site);

            this.RecordSchemaLock(fileUrl, clientId, schemaLockId, userName, password, domain);
        }

        /// <summary>
        /// This method is used to record the schema lock status with specified client ID and schema lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the schema lock.</param>
        /// <param name="clientId">Specify the client ID of the schema lock.</param>
        /// <param name="schemaLockId">Specify the schema ID of the schema lock.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        public void RecordSchemaLock(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain)
        {
            Action function = () =>
            {
                SchemaLockSubRequestType schemaLockRequest = new SchemaLockSubRequestType();
                schemaLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
                schemaLockRequest.SubRequestData = new SchemaLockSubRequestDataType();
                schemaLockRequest.SubRequestData.SchemaLockRequestType = SchemaLockRequestTypes.ReleaseLock;
                schemaLockRequest.SubRequestData.SchemaLockRequestTypeSpecified = true;
                schemaLockRequest.SubRequestData.SchemaLockID = schemaLockId;
                schemaLockRequest.SubRequestData.ClientID = clientId;

                IMS_FSSHTTP_FSSHTTPBAdapter adapter = site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
                this.intializeContext(fileUrl, userName, password, domain);

                SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(adapter.CellStorageRequest(fileUrl, new SubRequestType[] { schemaLockRequest }), 0, 0, site);

                if (!string.Equals("Success", subResponse.ErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    this.errorMessage.Add(string.Format("Failed to release the schema lock for the client id {1} and schema lock id {5}on the file {0} using the following user: {2}/{3}and password:{4}", fileUrl, clientId, userName, domain, password, schemaLockId));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(new SharedLockKey(fileUrl, clientId, schemaLockId), function);
        }

        /// <summary>
        /// This method is used to record the coauth lock status with specified client ID and schema lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the coauth lock.</param>
        /// <param name="clientId">Specify the client ID of the coauth lock.</param>
        /// <param name="schemaLockId">Specify the schema ID of the coauth lock.</param>
        public void RecordCoauthSession(string fileUrl, string clientId, string schemaLockId)
        {
            string userName = Common.GetConfigurationPropertyValue("UserName1", this.site);
            string password = Common.GetConfigurationPropertyValue("Password1", this.site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.site);

            this.RecordCoauthSession(fileUrl, clientId, schemaLockId, userName, password, domain);
        }

        /// <summary>
        /// This method is used to record the coauth lock status with specified client ID and schema lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the coauth lock.</param>
        /// <param name="clientId">Specify the client ID of the coauth lock.</param>
        /// <param name="schemaLockId">Specify the schema ID of the coauth lock.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        public void RecordCoauthSession(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain)
        {
            Action function = () =>
            {
                CoauthSubRequestType coauthLockRequest = new CoauthSubRequestType();
                coauthLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
                coauthLockRequest.SubRequestData = new CoauthSubRequestDataType();
                coauthLockRequest.SubRequestData.CoauthRequestType = CoauthRequestTypes.ExitCoauthoring;
                coauthLockRequest.SubRequestData.CoauthRequestTypeSpecified = true;
                coauthLockRequest.SubRequestData.SchemaLockID = schemaLockId;
                coauthLockRequest.SubRequestData.ClientID = clientId;

                IMS_FSSHTTP_FSSHTTPBAdapter adapter = site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
                this.intializeContext(fileUrl, userName, password, domain);

                CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(adapter.CellStorageRequest(fileUrl, new SubRequestType[] { coauthLockRequest }), 0, 0, site);

                if (!string.Equals("Success", subResponse.ErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    this.errorMessage.Add(string.Format("Failed to leave the coauth session for the client id {1} and schema lock id {5} on the file {0} using the following user: {2}/{3}and password:{4}", fileUrl, clientId, userName, domain, password, schemaLockId));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(new SharedLockKey(fileUrl, clientId, schemaLockId), function);
        }

        /// <summary>
        /// This method is used to record the exclusive lock status with specified exclusive lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the exclusive lock.</param>
        /// <param name="exclusiveId">Specify the exclusive lock ID of the exclusive lock.</param>
        public void RecordExclusiveLock(string fileUrl, string exclusiveId)
        {
            string userName = Common.GetConfigurationPropertyValue("UserName1", this.site);
            string password = Common.GetConfigurationPropertyValue("Password1", this.site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.site);

            this.RecordExclusiveLock(fileUrl, exclusiveId, userName, password, domain);
        }

        /// <summary>
        /// This method is used to record the exclusive lock status with specified exclusive lock ID for the file URL.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the exclusive lock.</param>
        /// <param name="exclusiveId">Specify the exclusive lock ID of the exclusive lock.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        public void RecordExclusiveLock(string fileUrl, string exclusiveId, string userName, string password, string domain)
        {
            Action function = () =>
            {
                ExclusiveLockSubRequestType coauthLockRequest = new ExclusiveLockSubRequestType();
                coauthLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
                coauthLockRequest.SubRequestData = new ExclusiveLockSubRequestDataType();
                coauthLockRequest.SubRequestData.ExclusiveLockRequestType = ExclusiveLockRequestTypes.ReleaseLock;
                coauthLockRequest.SubRequestData.ExclusiveLockRequestTypeSpecified = true;
                coauthLockRequest.SubRequestData.ExclusiveLockID = exclusiveId;

                IMS_FSSHTTP_FSSHTTPBAdapter adapter = site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();
                this.intializeContext(fileUrl, userName, password, domain);

                CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(adapter.CellStorageRequest(fileUrl, new SubRequestType[] { coauthLockRequest }), 0, 0, site);

                if (!string.Equals("Success", subResponse.ErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    this.errorMessage.Add(string.Format("Failed to release the exclusive lock for the exclusive lock id {1} on the file {0} using the following user: {2}/{3}and password:{4}", fileUrl, exclusiveId, userName, domain, password));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(fileUrl, function, KeyStatus.ExclusiveLock);
        }

        /// <summary>
        /// This method is used to record the status of saving file to document library that needs files checked out.
        /// </summary>
        public void RecordDocumentLibraryCheckOutRequired()
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.ChangeDocLibraryStatus(false))
                {
                    this.errorMessage.Add("Failed to change the status of saving document to the library that does not require check out files.");
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(ServerStatus.CheckOutRequired, function);
        }

        /// <summary>
        /// This method is used to record the status of storage web service turned off.
        /// </summary>
        public void RecordDisableCellStorageService()
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.SwitchCellStorageService(true))
                {
                    this.errorMessage.Add("Failed to enable coauthoring feature.");
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(ServerStatus.DisableCellStorageWebService, function);
        }

        /// <summary>
        /// This method is used to record the status of disable the versioning for the specified document library.
        /// </summary>
        /// <param name="documentLibraryName">Specify the document library name.</param>
        public void RecordDisableVersioning(string documentLibraryName)
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.SwitchMajorVersioning(documentLibraryName, true))
                {
                    this.errorMessage.Add(string.Format("Failed to enable versioning on the document library {0}.", documentLibraryName));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(ServerStatus.DisableVersioning, function);
        }

        /// <summary>
        /// This method is used to record the specified file URL which has been checked out.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which has been checked out.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        public void RecordFileCheckOut(string fileUrl, string userName, string password, string domain)
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.CheckInFile(fileUrl, userName, password, domain, "Check In for test purpose."))
                {
                    this.errorMessage.Add(string.Format("Failed to check in the file {0} using the user {1}/{2} and password: {3}", fileUrl, userName, domain, password));
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(fileUrl, function, KeyStatus.CheckOut);
        }

        /// <summary>
        /// This method is used to record the status which one file is uploaded to the specified full file URI.
        /// </summary>
        /// <param name="url">Specify the full file URL.</param>
        public void RecordFileUpload(string url)
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter>();
                string fileUrl = url.Substring(0, url.LastIndexOf("/", StringComparison.OrdinalIgnoreCase));
                string fileName = url.Substring(url.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);
                try
                {
                    if (!sutAdapter.RemoveFile(fileUrl, fileName))
                    {
                        this.errorMessage.Add(string.Format("Cannot remove a file in the URL {0}", url));
                        this.isEnvironmentRollbackSuccess = false;
                    }
                }
                catch (Microsoft.VisualStudio.TestTools.UnitTesting.AssertInconclusiveException e)
                {
                    // Here try to catch the exception to avoid the removing file failure due to the server cannot release the lock as expected.
                    this.errorMessage.Add(e.Message);
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(url, function, KeyStatus.UploadTextFile);
        }

        /// <summary>
        /// This method is used to cancel record of coauthoring session or schema lock. 
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the coauth lock.</param>
        /// <param name="clientId">Specify the client ID of the coauth lock.</param>
        /// <param name="schemaLockId">Specify the schema ID of the coauth lock.</param>
        public void CancelSharedLock(string fileUrl, string clientId, string schemaLockId)
        {
            SharedLockKey key = new SharedLockKey(fileUrl, clientId, schemaLockId);
            if (this.releaseSharedLockFunctions.Keys.Contains(key))
            {
                this.releaseSharedLockFunctions.Remove(key);
            }
        }

        /// <summary>
        /// This method is used to cancel record of exclusive lock. 
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which get the exclusive lock.</param>
        public void CancelExclusiveLock(string fileUrl)
        {
            if (this.releaseExclusiveLockFunctions.Keys.Contains(fileUrl))
            {
                this.releaseExclusiveLockFunctions.Remove(fileUrl);
            }
        }

        /// <summary>
        /// This method is used to cancel record of check out.
        /// </summary>
        /// <param name="fileUrl">Specify the file URL which has been checked out.</param>
        public void CancelRecordCheckOut(string fileUrl)
        {
            if (this.fileCheckOutRollbackFunctions.Keys.Contains(fileUrl))
            {
                this.fileCheckOutRollbackFunctions.Remove(fileUrl);
            }
        }

        /// <summary>
        /// This method is used to record the status of disable coauthoring feature.
        /// </summary>
        public void RecordDisableCoauth()
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.SwitchCoauthoringFeature(false))
                {
                    this.errorMessage.Add("Failed to enable coauthoring feature.");
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(ServerStatus.DisableCoauth, function);
        }

        /// <summary>
        /// This method is used to record the status of switching to the windows based claims authentication.
        /// </summary>
        public void RecordDisableClaimsBasedAuthentication()
        {
            Action function = () =>
            {
                IMS_FSSHTTP_FSSHTTPBSUTControlAdapter sutAdapter = this.site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
                if (!sutAdapter.SwitchClaimsAuthentication(true))
                {
                    this.errorMessage.Add("Failed to enable claims based authentication.");
                    this.isEnvironmentRollbackSuccess = false;
                }
            };

            this.AddOrUpdate(ServerStatus.DisableClaimsBasedAuthentication, function);
        }

        /// <summary>
        /// This method is used to add or update the ReleaseSharedLockFunctions key/value pairs.
        /// </summary>
        /// <param name="key">Specify the key.</param>
        /// <param name="value">Specify the value.</param>
        private void AddOrUpdate(SharedLockKey key, Action value)
        {
            if (this.releaseSharedLockFunctions.Keys.Contains(key))
            {
                this.releaseSharedLockFunctions[key] = value;
            }
            else
            {
                // If the file already lock by the exclusive lock, then remove the exclusive lock record.
                string findKey = this.releaseExclusiveLockFunctions.Keys.FirstOrDefault(fileUrl => fileUrl == key.FileUrl);
                if (findKey != null)
                {
                    this.releaseExclusiveLockFunctions.Remove(findKey);
                }

                this.releaseSharedLockFunctions.Add(key, value);
            }
        }

        /// <summary>
        /// This method is used to add or update the DocumentLibraryStatusRollbackFunctions key/value pairs.
        /// </summary>
        /// <param name="key">Specify the key.</param>
        /// <param name="value">Specify the value.</param>
        private void AddOrUpdate(ServerStatus key, Action value)
        {
            if (this.documentLibraryStatusRollbackFunctions.Keys.Contains(key))
            {
                this.documentLibraryStatusRollbackFunctions[key] = value;
            }
            else
            {
                this.documentLibraryStatusRollbackFunctions.Add(key, value);
            }
        }

        /// <summary>
        /// This method is used to add or update the key/value pairs with the specified key status.
        /// </summary>
        /// <param name="key">Specify the key.</param>
        /// <param name="value">Specify the value.</param>
        /// <param name="keyStatus">Specify the key status.</param>
        private void AddOrUpdate(string key, Action value, KeyStatus keyStatus)
        {
            switch (keyStatus)
            {
                case KeyStatus.CheckOut:
                    if (this.fileCheckOutRollbackFunctions.Keys.Contains(key))
                    {
                        this.fileCheckOutRollbackFunctions[key] = value;
                    }
                    else
                    {
                        this.fileCheckOutRollbackFunctions.Add(key, value);
                    }

                    break;

                case KeyStatus.ExclusiveLock:
                    if (this.releaseExclusiveLockFunctions.Keys.Contains(key))
                    {
                        // There is no way to record twice exclusive lock for just one file
                        this.site.Assert.Fail("Fail to record the exclusive lock on the file {0} more than once.", key);
                    }
                    else
                    {
                        // If the file already lock by the shared lock, then remove the shared lock record.
                        SharedLockKey findKey = this.releaseSharedLockFunctions.Keys.FirstOrDefault(k => k.FileUrl == key);
                        if (findKey != null)
                        {
                            this.releaseSharedLockFunctions.Remove(findKey);
                        }

                        this.releaseExclusiveLockFunctions.Add(key, value);
                    }

                    break;

                case KeyStatus.UploadTextFile:
                    if (this.removeFileFunctions.Keys.Contains(key))
                    {
                        this.removeFileFunctions[key] = value;
                    }
                    else
                    {
                        this.removeFileFunctions.Add(key, value);
                    }

                    break;

                default:
                    this.site.Assert.Fail("Unsupported operation.");
                    break;
            }
        }

        /// <summary>
        /// The class is used for the key when recording the coauthoring session or schema lock.
        /// </summary>
        public class SharedLockKey
        {
            /// <summary>
            /// Initializes a new instance of the SharedLockKey class.
            /// </summary>
            /// <param name="fileUrl">Specify the file URL.</param>
            /// <param name="clientId">Specify the client id.</param>
            /// <param name="schemalockId">Specify the schema lock id.</param>
            public SharedLockKey(string fileUrl, string clientId, string schemalockId)
            {
                this.FileUrl = fileUrl;
                this.ClientId = clientId;
                this.SchemaLockId = schemalockId;
            }

            /// <summary>
            /// Gets or sets the file URL.
            /// </summary>
            public string FileUrl { get; set; }

            /// <summary>
            /// Gets or sets the client id.
            /// </summary>
            public string ClientId { get; set; }

            /// <summary>
            /// Gets or sets the schema lock id.
            /// </summary>
            public string SchemaLockId { get; set; }

            /// <summary>
            /// Override the equals function.
            /// </summary>
            /// <param name="obj">Specify the other instance need to be compared.</param>
            /// <returns>Return true if equals, otherwise return false.</returns>
            public override bool Equals(object obj)
            {
                if (obj == null)
                {
                    return false;
                }

                if (this.GetType() == obj.GetType())
                {
                    SharedLockKey other = (SharedLockKey)obj;
                    return this.ClientId == other.ClientId && this.FileUrl == other.FileUrl && this.SchemaLockId == other.SchemaLockId;
                }
                else
                {
                    return false;
                }
            }

            /// <summary>
            /// Override the GetHashCode function.
            /// </summary>
            /// <returns>Return the hash code.</returns>
            public override int GetHashCode()
            {
                return this.FileUrl.GetHashCode() + this.ClientId.GetHashCode() + this.SchemaLockId.GetHashCode();
            }
        }
    }
}