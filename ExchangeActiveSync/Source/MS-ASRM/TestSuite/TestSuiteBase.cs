namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using DataStructures=Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// A bass class for scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets or sets the related information of User1.
        /// </summary>
        protected UserInformation UserOneInformation { get; set; }

        /// <summary>
        /// Gets or sets the related information of User2.
        /// </summary>
        protected UserInformation UserTwoInformation { get; set; }

        /// <summary>
        /// Gets or sets the related information of User3.
        /// </summary>
        protected UserInformation UserThreeInformation { get; set; }

        /// <summary>
        /// Gets or sets the related information of User4.
        /// </summary>
        protected UserInformation UserFourInformation { get; set; }

        /// <summary>
        /// Gets the MS-ASRM protocol adapter.
        /// </summary>
        protected IMS_ASRMAdapter ASRMAdapter { get; private set; }

        /// <summary>
        /// Gets MS-ASRM SUT Control adapter.
        /// </summary>
        protected IMS_ASRMSUTControlAdapter ASRMSUTControlAdapter { get; private set; }
        #endregion

        /// <summary>
        /// Add the item created in test case to the item collections needed to clean up.
        /// </summary>
        /// <param name="userInformation">The identity of the user who has the item.</param>
        /// <param name="parentFolder">The CollectionId of the folder in which the item is placed.</param>
        /// <param name="itemSubject">The subject of the item to delete.</param>
        protected static void AddCreatedItemToCollection(UserInformation userInformation, string parentFolder, string itemSubject)
        {
            CreatedItems createdItems = new CreatedItems { CollectionId = parentFolder };
            createdItems.ItemSubject.Add(itemSubject);
            userInformation.UserCreatedItems.Add(createdItems);
        }

        #region Test case initialize and cleanup
        /// <summary>
        /// Override the base TestInitialize function
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();

            if (this.ASRMAdapter == null)
            {
                this.ASRMAdapter = this.Site.GetAdapter<IMS_ASRMAdapter>();
            }

            this.ASRMSUTControlAdapter = this.Site.GetAdapter<IMS_ASRMSUTControlAdapter>();

            // If implementation doesn't support this specification [MS-ASRM], the case will not start.
            if (!bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", this.Site)))
            {
                this.Site.Assert.Inconclusive("This test suite is not supported under current SUT, because MS-ASRM_Supported value is set to false in MS-ASRM_{0}_SHOULDMAY.deployment.ptfconfig file.", Common.GetSutVersion(this.Site));
            }

            // Set the information of User1.
            this.UserOneInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Set the information of User2.
            this.UserTwoInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Set the information of User3.
            this.UserThreeInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User3Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User3Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };

            // Set the information of User4.
            this.UserFourInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User4Name", this.Site),
                UserPassword = Common.GetConfigurationPropertyValue("User4Password", this.Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", this.Site)
            };
            string sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);

            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", this.Site)))
            {
                if (Common.GetConfigurationPropertyValue("TransportType", this.Site)
                    .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
                {
                    // Use the user who is in Administrators group to enable the SSL setting.
                    bool isSSLUpdated = this.ASRMSUTControlAdapter.ConfigureSSLSetting(
                        sutComputerName,
                        this.UserFourInformation.UserName,
                        this.UserFourInformation.UserPassword,
                        this.UserFourInformation.UserDomain,
                        true);
                    this.Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be enabled.");
                }
            }

            if (Common.GetSutVersion(this.Site) != SutVersion.ExchangeServer2007 || string.Equals(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site), "12.1"))
            {
                this.SwitchUser(this.UserOneInformation, true);
            }
        }

        /// <summary>
        /// Override the base TestCleanup function
        /// </summary>
        protected override void TestCleanup()
        {
            // If implementation doesn't support this specification [MS-ASRM], the case will not start.
            if (bool.Parse(Common.GetConfigurationPropertyValue("MS-ASRM_Supported", this.Site)))
            {
                string sutComputerName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);

                if (Common.GetConfigurationPropertyValue("TransportType", this.Site)
                    .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
                {
                    // Use the user who is in Administrators group to disable the SSL setting.
                    bool isSSLUpdated = this.ASRMSUTControlAdapter.ConfigureSSLSetting(
                        sutComputerName,
                        this.UserFourInformation.UserName,
                        this.UserFourInformation.UserPassword,
                        this.UserFourInformation.UserDomain,
                        false);
                    this.Site.Assert.IsTrue(isSSLUpdated, "The SSL setting of protocol web service should be disabled.");
                }

                // Clean the created items of User1.
                if (this.UserOneInformation.UserCreatedItems.Count != 0)
                {
                    // Switch the user credential to User1.
                    this.SwitchUser(this.UserOneInformation, false);
                    this.DeleteCreatedItems(this.UserOneInformation.UserCreatedItems);
                }

                if (this.UserTwoInformation.UserCreatedItems.Count != 0)
                {
                    // Switch the user credential to User2.
                    this.SwitchUser(this.UserTwoInformation, false);
                    this.DeleteCreatedItems(this.UserTwoInformation.UserCreatedItems);
                }

                if (this.UserThreeInformation.UserCreatedItems.Count != 0)
                {
                    // Switch the user credential to User3.
                    this.SwitchUser(this.UserThreeInformation, false);
                    this.DeleteCreatedItems(this.UserThreeInformation.UserCreatedItems);
                }
            }

            base.TestCleanup();
        }

        #endregion

        #region protected methods
        /// <summary>
        /// Checks if ActiveSync Protocol Version is "14.1" and Transport Type is "HTTPS".
        /// </summary>
        protected void CheckPreconditions()
        {
            Site.Assume.IsTrue(Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1") ||
                                Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0") ||
                                Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"),
                                "Implementation does consider the XML body of the command request to be invalid, if the protocol version specified by in the command request is not 14.1 and 16.0.");
            Site.Assume.AreEqual<string>("HTTPS", Common.GetConfigurationPropertyValue("TransportType", this.Site).ToUpper(CultureInfo.CurrentCulture), "This protocol requires that communication between the client and server occurs over an HTTP connection that uses Secure Sockets Layer (SSL).");
        }

        /// <summary>
        /// Call Settings command to get the expected template ID for template name
        /// </summary>
        /// <param name="templateName">A string that specifies the name of the rights policy template.</param>
        /// <returns>A string that identifies a particular rights policy template to be applied to the outgoing message.</returns>
        protected string GetTemplateID(string templateName)
        {
            // Get the template settings
            SettingsResponse settingsResponse = this.ASRMAdapter.Settings();

            // Choose the all rights policy template and get template ID.
            this.Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation, "The RightsManagementInformation element should not be null in Settings response.");
            this.Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Get, "The Get element should not be null in Settings response.");
            this.Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Get.RightsManagementTemplates, "The RightsManagementTemplates element should not be null in Settings response.");
            string templateID = null;
            foreach (Response.RightsManagementTemplatesRightsManagementTemplate template in settingsResponse.ResponseData.RightsManagementInformation.Get.RightsManagementTemplates)
            {
                if (template.TemplateName == templateName)
                {
                    templateID = template.TemplateID;
                    break;
                }
            }

            this.Site.Assert.IsNotNull(templateID, "Template {0} is not found on the server. This may happen if MS-ASRM configuration is not performed properly.", templateName);
            return templateID;
        }

        /// <summary>
        /// Find an email with specific subject and folder.
        /// </summary>
        /// <param name="subject">The subject of the email item.</param>
        /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
        /// <param name="rightsManagementSupport">A boolean value specifies whether the server will decompress and decrypt rights-managed email messages before sending them to the client or not</param>
        /// <param name="isRetryNeeded">A boolean value specifies whether need retry.</param>
        /// <returns>Return change result</returns>
        protected DataStructures.Sync SyncEmail(string subject, string collectionId, bool? rightsManagementSupport, bool isRetryNeeded)
        {
            SyncRequest syncRequest = Common.CreateInitialSyncRequest(collectionId);
            DataStructures.SyncStore initSyncResult = this.ASRMAdapter.Sync(syncRequest);

            // Verify sync change result
            this.Site.Assert.AreEqual<byte>(1, initSyncResult.CollectionStatus, "If the Sync command executes successfully, the Status in response should be 1.");

            syncRequest = TestSuiteHelper.CreateSyncRequest(initSyncResult.SyncKey, collectionId, rightsManagementSupport);
            DataStructures.Sync sync = this.ASRMAdapter.SyncEmail(syncRequest, subject, isRetryNeeded);
            return sync;
        }

        /// <summary>
        /// Sync changes between client and server
        /// </summary>
        /// <param name="syncKey">The synchronization key returned by last request.</param>
        /// <param name="collectionId">Identify the folder as the collection being synchronized.</param>
        /// <param name="rightsManagementSupport">A boolean value specifies whether the server will decompress and decrypt rights-managed email messages before sending them to the client or not</param>
        /// <returns>Return change result</returns>
        protected DataStructures.SyncStore SyncChanges(string syncKey, string collectionId, bool rightsManagementSupport)
        {
            // Get changes from server use initial syncKey
            SyncRequest syncRequest = TestSuiteHelper.CreateSyncRequest(syncKey, collectionId, rightsManagementSupport);
            DataStructures.SyncStore syncResult = this.ASRMAdapter.Sync(syncRequest);

            return syncResult;
        }

        /// <summary>
        /// Change user to call FolderSync command to synchronize the collection hierarchy.
        /// </summary>
        /// <param name="userInformation">The information of a user.</param>
        /// <param name="syncFolderHierarchy">Whether sync folder hierarchy or not.</param>
        protected void SwitchUser(UserInformation userInformation, bool syncFolderHierarchy)
        {
            this.ASRMAdapter.SwitchUser(userInformation.UserName, userInformation.UserPassword, userInformation.UserDomain);

            if (syncFolderHierarchy)
            {
                FolderSyncResponse folderSyncResponse = this.ASRMAdapter.FolderSync(Common.CreateFolderSyncRequest("0"));

                // Get the CollectionId from FolderSync command response.
                if (string.IsNullOrEmpty(userInformation.InboxCollectionId))
                {
                    userInformation.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.SentItemsCollectionId))
                {
                    userInformation.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                }
            }
        }

        /// <summary>
        /// User1 sends mail to User2 and does FolderSync in User2's mailbox
        /// </summary>
        /// <param name="templateID">A string that identifies a particular rights policy template to be applied to the outgoing message.</param>
        /// <param name="saveInSentItems">A boolean that represent to save the sent item in SentItems folder or not.</param>
        /// <param name="copyToUserInformation">The UserInformation for the Cc element.</param>
        /// <returns>The subject of the sent message</returns>
        protected string SendMailAndFolderSync(string templateID, bool saveInSentItems, UserInformation copyToUserInformation)
        {
            #region User1 calls method SendMail to send MIME-formatted e-mail messages to User2
            string subject = Common.GenerateResourceName(this.Site, "subject");
            string content = "The content of the body.";
            SendMailRequest sendMailRequest;

            if (copyToUserInformation != null)
            {
                sendMailRequest = TestSuiteHelper.CreateSendMailRequest(Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain), Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain), Common.GetMailAddress(copyToUserInformation.UserName, copyToUserInformation.UserDomain), string.Empty, subject, content, templateID);
            }
            else
            {
                sendMailRequest = TestSuiteHelper.CreateSendMailRequest(Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain), Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain), string.Empty, string.Empty, subject, content, templateID);
            }

            if (saveInSentItems)
            {
                sendMailRequest.RequestData.SaveInSentItems = string.Empty;
            }

            SendMailResponse sendMailResponse = this.ASRMAdapter.SendMail(sendMailRequest);

            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("SSLWaitTime", this.Site));
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("SSLRetryCount", this.Site));

            if (Common.GetConfigurationPropertyValue("TransportType", this.Site)
                .Equals("HTTPS", StringComparison.CurrentCultureIgnoreCase))
            {
                while (!string.IsNullOrEmpty(sendMailResponse.ResponseDataXML) && counter < upperBound)
                {
                    // Await the SSL configuration take effect.
                    System.Threading.Thread.Sleep(waitTime);
                    sendMailResponse = this.ASRMAdapter.SendMail(sendMailRequest);
                    counter++;
                }
            }

            this.Site.Assert.AreEqual<string>(string.Empty, sendMailResponse.ResponseDataXML, "If the message is sent successfully, the server should return an empty response.");

            if (saveInSentItems)
            {
                AddCreatedItemToCollection(this.UserOneInformation, this.UserOneInformation.SentItemsCollectionId, subject);
            }
            #endregion

            #region Record Cc user name, folder collectionId and item subject that are used in case
            if (copyToUserInformation != null)
            {
                this.SwitchUser(copyToUserInformation, true);
                AddCreatedItemToCollection(copyToUserInformation, copyToUserInformation.InboxCollectionId, subject);
            }

            #endregion

            #region User2 calls method FolderSync to synchronize the collection hierarchy, returns folder collectionIds
            this.SwitchUser(this.UserTwoInformation, true);
            #endregion

            #region Record User2's user name, folder collectionId and item subject that are used in case
            AddCreatedItemToCollection(this.UserTwoInformation, this.UserTwoInformation.InboxCollectionId, subject);
            #endregion

            return subject;
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Delete the specified item.
        /// </summary>
        /// <param name="itemsToDelete">The collection of the items to delete.</param>
        private void DeleteCreatedItems(Collection<CreatedItems> itemsToDelete)
        {
            foreach (CreatedItems itemToDelete in itemsToDelete)
            {
                SyncRequest syncRequest = Common.CreateInitialSyncRequest(itemToDelete.CollectionId);
                DataStructures.SyncStore initSyncResult = this.ASRMAdapter.Sync(syncRequest);
                DataStructures.SyncStore result = this.SyncChanges(initSyncResult.SyncKey, itemToDelete.CollectionId, false);
                int i = 0;
                if (result.AddElements != null)
                {
                    Request.SyncCollectionDelete[] deletes = new Request.SyncCollectionDelete[result.AddElements.Count];
                    foreach (DataStructures.Sync item in result.AddElements)
                    {
                        foreach (string subject in itemToDelete.ItemSubject)
                        {
                            if (item.Email.Subject.Equals(subject))
                            {
                                Request.SyncCollectionDelete delete = new Request.SyncCollectionDelete
                                {
                                    ServerId = item.ServerId
                                };
                                deletes[i] = delete;
                            }
                        }

                        i++;
                    }

                    Request.SyncCollection syncCollection = TestSuiteHelper.CreateSyncCollection(result.SyncKey, itemToDelete.CollectionId);
                    syncCollection.Commands = deletes;

                    syncRequest = Common.CreateSyncRequest(new Request.SyncCollection[] { syncCollection });
                    DataStructures.SyncStore deleteResult = this.ASRMAdapter.Sync(syncRequest);
                    this.Site.Assert.AreEqual<byte>(
                        1,
                        deleteResult.CollectionStatus,
                        "The value of 'Status' should be 1 which indicates the Sync command executes successfully.");
                }
            }
        }

        #endregion
    }
}