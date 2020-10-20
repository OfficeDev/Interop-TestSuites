namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Gets MS-ASHTTP protocol adapter.
        /// </summary>
        protected IMS_ASHTTPAdapter HTTPAdapter { get; private set; }

        /// <summary>
        /// Gets MS-ASHTTP SUT Control adapter.
        /// </summary>
        protected IMS_ASHTTPSUTControlAdapter HTTPSUTControlAdapter { get; private set; }

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
        /// Gets or sets the related information of an invalid user.
        /// </summary>
        protected UserInformation InvalidUserInformation { get; set; }
        #endregion

        #region Test case initialize and cleanup
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.HTTPAdapter = Site.GetAdapter<IMS_ASHTTPAdapter>();
            this.HTTPSUTControlAdapter = Site.GetAdapter<IMS_ASHTTPSUTControlAdapter>();

            // Set the information of User1.
            this.UserOneInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User1Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            // Set the information of User2.
            this.UserTwoInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User2Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User2Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            // Set the information of User3.
            this.UserThreeInformation = new UserInformation
            {
                UserName = Common.GetConfigurationPropertyValue("User3Name", Site),
                UserPassword = Common.GetConfigurationPropertyValue("User3Password", Site),
                UserDomain = Common.GetConfigurationPropertyValue("Domain", Site)
            };

            // Initialize the UserInformation of an invalid user.
            this.InvalidUserInformation = new UserInformation();

            // Initialize the HTTP POST request prefix field.
            this.InitializeRequestPrefixField();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            // Clean the created items of User1.
            if (this.UserOneInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User1.
                this.SwitchUser(this.UserOneInformation, true);

                this.DeleteCreatedItems(this.UserOneInformation.UserCreatedItems);
            }

            // Clean the created items of User2.
            if (this.UserTwoInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User2.
                this.SwitchUser(this.UserTwoInformation, true);

                this.DeleteCreatedItems(this.UserTwoInformation.UserCreatedItems);
            }

            // Clean the created items of User3.
            if (this.UserThreeInformation.UserCreatedItems.Count != 0)
            {
                // Switch the user credential to User3.
                this.SwitchUser(this.UserThreeInformation, true);

                this.DeleteCreatedItems(this.UserThreeInformation.UserCreatedItems);
            }

            // Reset the user credential.
            this.SwitchUser(this.UserOneInformation, false);
            base.TestCleanup();
        }
        #endregion

        #region Test case base methods
        /// <summary>
        /// Change the user authentication.
        /// </summary>
        /// <param name="userInformation">The information of a user.</param>
        /// <param name="syncFolderHierarchy">Whether sync folder hierarchy or not.</param>
        protected void SwitchUser(UserInformation userInformation, bool syncFolderHierarchy)
        {
            Dictionary<HTTPPOSTRequestPrefixField, string> requestPrefixFields = new Dictionary<HTTPPOSTRequestPrefixField, string>
            {
                {
                    HTTPPOSTRequestPrefixField.UserName, userInformation.UserName
                },
                {
                    HTTPPOSTRequestPrefixField.Password, userInformation.UserPassword
                },
                {
                    HTTPPOSTRequestPrefixField.Domain, userInformation.UserDomain
                }
            };

            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefixFields);

            if (syncFolderHierarchy)
            {
                // Call FolderSync command to synchronize the collection hierarchy.
                FolderSyncResponse folderSyncResponse = this.CallFolderSyncCommand();

                // Get the CollectionId from FolderSync command response.
                if (string.IsNullOrEmpty(userInformation.InboxCollectionId))
                {
                    userInformation.InboxCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Inbox, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.SentItemsCollectionId))
                {
                    userInformation.SentItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.DeletedItemsCollectionId))
                {
                    userInformation.DeletedItemsCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.DeletedItems, this.Site);
                }

                if (string.IsNullOrEmpty(userInformation.CalendarCollectionId))
                {
                    userInformation.CalendarCollectionId = Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.Calendar, this.Site);
                }
            }
        }

        /// <summary>
        /// Call FolderSync command and get the response.
        /// </summary>
        /// <returns>The response of FolderSync command.</returns>
        protected FolderSyncResponse CallFolderSyncCommand()
        {
            // Create FolderSync command request.
            string folderSyncRequest = Common.CreateFolderSyncRequest("0").GetRequestDataSerializedXML();

            // Call FolderSync command by HTTP POST.
            SendStringResponse folderSyncResponseString = this.HTTPAdapter.HTTPPOST(CommandName.FolderSync, null, folderSyncRequest);

            // Check the command is executed successfully.
            this.CheckResponseStatus(folderSyncResponseString.ResponseDataXML);

            // Convert SendStringResponse to FolderSyncResponse.
            FolderSyncResponse folderSyncResponse = new FolderSyncResponse
            {
                ResponseDataXML = folderSyncResponseString.ResponseDataXML,
                Headers = folderSyncResponseString.Headers
            };

            folderSyncResponse.DeserializeResponseData();

            return folderSyncResponse;
        }

        /// <summary>
        /// Call SendMail command to send mail to the server.
        /// </summary>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="to">The mailbox address of recipient.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="saveInSent">The value of SaveInSent command parameter.</param>
        protected void CallSendMailCommand(string from, string to, string subject, string saveInSent)
        {
            // Create the SendMail command request.
            string mime = Common.CreatePlainTextMime(from, to, string.Empty, string.Empty, subject, "SendMail content");
            SendMailRequest request = Common.CreateSendMailRequest(null, System.Guid.NewGuid().ToString(), mime);

            // If saveInSent is not null, set the SaveInSent command parameter.
            if (saveInSent != null)
            {
                request.SetCommandParameters(new Dictionary<CmdParameterName, object>());
                int result;

                // If saveInSent is "0" or "1", the Options command parameter should be set.
                // If saveInSent is "T" or "F", the SaveInSent command parameter should be set.
                if (int.TryParse(saveInSent, out result))
                {
                    request.CommandParameters.Add(CmdParameterName.Options, result);
                    if (result == 1)
                    {
                        request.RequestData.SaveInSentItems = string.Empty;
                    }
                }
                else
                {
                    request.CommandParameters.Add(CmdParameterName.SaveInSent, saveInSent);
                    if (saveInSent == "T")
                    {
                        request.RequestData.SaveInSentItems = string.Empty;
                    }
                }
            }

            // Call SendMail command by HTTP POST.
            SendStringResponse sendMailResponse = this.HTTPAdapter.HTTPPOST(CommandName.SendMail, request.CommandParameters, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(sendMailResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call SmartForward command to forward an email.
        /// </summary>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="forwardTo">The mailbox address of recipient.</param>
        /// <param name="itemServerId">The ServerId of the email to forward.</param>
        /// <param name="forwardSubject">The subject of the email to forward.</param>
        /// <param name="saveInSent">The value of SaveInSent command parameter.</param>
        /// <param name="longId">The value of LongId command parameter.</param>
        /// <param name="occurrence">The value of Occurrence command parameter.</param>
        protected void CallSmartForwardCommand(string from, string forwardTo, string itemServerId, string forwardSubject, string saveInSent, string longId, string occurrence)
        {
            // Create SmartForward command request.
            Request.Source source = new Request.Source();

            string mime = Common.CreatePlainTextMime(from, forwardTo, string.Empty, string.Empty, forwardSubject, "SmartForward content");
            SmartForwardRequest request = Common.CreateSmartForwardRequest(null, System.Guid.NewGuid().ToString(), mime, source);

            // Set the command parameters.
            // If the LongId element is present, the FolderId, ItemId, and InstanceId elements are not present.
            request.SetCommandParameters(new Dictionary<CmdParameterName, object>());
            if (longId == null)
            {
                string collectionId = itemServerId.Split(':')[0];
                if (saveInSent != null)
                {
                    int result;
                    if (int.TryParse(saveInSent, out result))
                    {
                        request.CommandParameters.Add(CmdParameterName.Options, result);
                        if (result == 1)
                        {
                            request.RequestData.SaveInSentItems = string.Empty;
                        }
                    }
                    else
                    {
                        request.CommandParameters.Add(CmdParameterName.SaveInSent, saveInSent);
                        if (saveInSent == "T")
                        {
                            request.RequestData.SaveInSentItems = string.Empty;
                        }
                    }
                }

                // If the InstanceId element is present, both the FolderId and ItemId elements SHOULD be present.
                if (occurrence != null)
                {
                    request.RequestData.Source.InstanceId = occurrence;
                    request.CommandParameters.Add(CmdParameterName.Occurrence, occurrence);
                }

                source.FolderId = collectionId;
                source.ItemId = itemServerId;
                request.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
                request.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
            }
            else
            {
                request.RequestData.Source.LongId = longId;
                request.CommandParameters.Add(CmdParameterName.LongId, longId);
            }

            // Call SmartForward command by HTTP POST.
            SendStringResponse smartForwardResponse = this.HTTPAdapter.HTTPPOST(CommandName.SmartForward, request.CommandParameters, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(smartForwardResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call SmartReply command to reply an email.
        /// </summary>
        /// <param name="from">The mailbox address of sender.</param>
        /// <param name="replyTo">The mailbox address recipient.</param>
        /// <param name="itemServerId">The ServerId of the email to forward.</param>
        /// <param name="replySubject">The subject of the email to reply.</param>
        /// <param name="saveInSent">The value of SaveInSent command parameter.</param>
        /// <param name="longId">The value of LongId command parameter.</param>
        /// <param name="occurrence">The value of Occurrence command parameter.</param>
        protected void CallSmartReplyCommand(string from, string replyTo, string itemServerId, string replySubject, string saveInSent, string longId, string occurrence)
        {
            // Create SmartReply command request.
            Request.Source source = new Request.Source();
           
            string mime = Common.CreatePlainTextMime(from, replyTo, string.Empty, string.Empty, replySubject, "SmartReply content");
            SmartReplyRequest request = Common.CreateSmartReplyRequest(null, System.Guid.NewGuid().ToString(), mime, source);

            // Set the command parameters.
            // If the LongId element is present, the FolderId, ItemId, and InstanceId elements are not present.
            request.SetCommandParameters(new Dictionary<CmdParameterName, object>());
            if (longId == null)
            {
                string collectionId = itemServerId.Split(':')[0];
                if (saveInSent != null)
                {
                    int result;
                    if (int.TryParse(saveInSent, out result))
                    {
                        request.CommandParameters.Add(CmdParameterName.Options, result);
                        if (result == 1)
                        {
                            request.RequestData.SaveInSentItems = string.Empty;
                        }
                    }
                    else
                    {
                        request.CommandParameters.Add(CmdParameterName.SaveInSent, saveInSent);
                        if (saveInSent == "T")
                        {
                            request.RequestData.SaveInSentItems = string.Empty;
                        }
                    }
                }

                // If the InstanceId element is present, both the FolderId and ItemId elements SHOULD be present.
                if (occurrence != null)
                {
                    request.RequestData.Source.InstanceId = occurrence;
                    request.CommandParameters.Add(CmdParameterName.Occurrence, occurrence);
                }

                source.FolderId = collectionId;
                source.ItemId = itemServerId;
                request.CommandParameters.Add(CmdParameterName.CollectionId, collectionId);
                request.CommandParameters.Add(CmdParameterName.ItemId, itemServerId);
            }
            else
            {
                request.RequestData.Source.LongId = longId;
                request.CommandParameters.Add(CmdParameterName.LongId, longId);
            }

            // Call SmartReply command by HTTP POST.
            SendStringResponse smartReplyResponse = this.HTTPAdapter.HTTPPOST(CommandName.SmartReply, request.CommandParameters, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(smartReplyResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call Sync command to synchronize changes in a folder between the client and the server.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to sync.</param>
        /// <returns>The SyncStore abstracted from the response of Sync command.</returns>
        protected SyncStore CallSyncCommand(string collectionId)
        {
            // Call initial Sync command.
            SyncResponse syncResponse = this.CallInitialSyncCommand(collectionId);

            Site.Assert.IsNotNull(TestSuiteHelper.LoadSyncResponse(syncResponse), "The Sync response should not be null.");

            SyncRequest syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, TestSuiteHelper.LoadSyncResponse(syncResponse).SyncKey);
            syncRequest.RequestData.WindowSize = "512";

            // If there is MoreAvailable tag in the Sync command response, synchronize again to continue getting items from the server.
            SendStringResponse syncResponseString;
            SyncStore syncStoreTotal = new SyncStore();
            do
            {
                // Get the latest SyncKey.
                syncRequest.RequestData.Collections[0].SyncKey = TestSuiteHelper.LoadSyncResponse(syncResponse).SyncKey;

                // Call Sync command by HTTP POST using the SyncKey returned from last sync.
                syncResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Sync, null, syncRequest.GetRequestDataSerializedXML());

                if (!string.IsNullOrEmpty(syncResponseString.ResponseDataXML))
                {
                    // Check the command is executed successfully.
                    this.CheckResponseStatus(syncResponseString.ResponseDataXML);

                    // Convert from SendStringResponse to SyncResponse.
                    syncResponse = TestSuiteHelper.ConvertSyncResponseFromSendString(syncResponseString);
                    SyncStore syncStore = TestSuiteHelper.LoadSyncResponse(syncResponse);
                    syncStoreTotal.SyncKey = syncStore.SyncKey;
                    for (int i = 0; i < syncStore.AddCommands.Count; i++)
                    {
                        syncStoreTotal.AddCommands.Add(syncStore.AddCommands[i]);
                    }
                }
            }
            while (syncResponseString.ResponseDataXML.Contains("<MoreAvailable />"));

            // Return the SyncStore abstracted from the syncResponse.
            return syncStoreTotal;
        }

        /// <summary>
        /// Call Sync command to synchronize changes in a folder between the client and the server with "0" SyncKey.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to sync.</param>
        /// <returns>The SyncResponse response.</returns>
        protected SyncResponse CallInitialSyncCommand(string collectionId)
        {
            SyncRequest syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, "0");

            // Call Sync command by HTTP POST.
            SendStringResponse syncResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Sync, null, syncRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(syncResponseString.ResponseDataXML);

            // Convert from SendStringResponse to SyncResponse.
            SyncResponse syncResponse = TestSuiteHelper.ConvertSyncResponseFromSendString(syncResponseString);

            return syncResponse;
        }

        /// <summary>
        /// Call Search command to find email.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to search.</param>
        /// <param name="freeText">The key words to search.</param>
        /// <returns>The response of Search command.</returns>
        protected SearchResponse CallSearchCommand(string collectionId, string freeText)
        {
            // Create Search command request.
            Request.SearchStore[] searchStores = new Request.SearchStore[1];
            searchStores[0] = new Request.SearchStore
            {
                Name = SearchName.Mailbox.ToString(),
                Query = new Request.queryType
                {
                    ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.And },
                    Items = new Request.queryType[] { new Request.queryType() }
                }
            };

            ((Request.queryType)searchStores[0].Query.Items[0]).ItemsElementName = new Request.ItemsChoiceType2[] { Request.ItemsChoiceType2.CollectionId, Request.ItemsChoiceType2.FreeText };
            ((Request.queryType)searchStores[0].Query.Items[0]).Items = new object[] { collectionId, freeText };

            searchStores[0].Options = new Request.Options1
            {
                Items = new object[] { string.Empty, "0-9", string.Empty },
                ItemsElementName =
                    new Request.ItemsChoiceType6[]
                    {
                        Request.ItemsChoiceType6.RebuildResults, Request.ItemsChoiceType6.Range,
                        Request.ItemsChoiceType6.DeepTraversal
                    }
            };

            SearchRequest searchRequest = Common.CreateSearchRequest(searchStores);

            // Call Search command by HTTP POST.
            string searchString = searchRequest.GetRequestDataSerializedXML();
            SendStringResponse searchResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Search, null, searchString);

            // Convert from SendStringResponse to SearchResponse.
            SearchResponse searchResponse = new SearchResponse { ResponseDataXML = searchResponseString.ResponseDataXML };
            searchResponse.DeserializeResponseData();

            return searchResponse;
        }

        /// <summary>
        /// Call MeetingResponse command to accept a meeting.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder in which the meeting request is placed.</param>
        /// <param name="itemServerId">The ServerId of the meeting request.</param>
        protected void CallMeetingResponseCommand(string collectionId, string itemServerId)
        {
            // Create MeetingResponse request.
            Request.MeetingResponseRequest[] meetingResponseRequest = new Request.MeetingResponseRequest[1];
            meetingResponseRequest[0] = new Request.MeetingResponseRequest
            {
                CollectionId = collectionId,
                RequestId = itemServerId,
                UserResponse = 1
            };
            MeetingResponseRequest request = Common.CreateMeetingResponseRequest(meetingResponseRequest);

            // Call MeetingResponse command by HTTP POST.
            SendStringResponse meetingResponseResponse = this.HTTPAdapter.HTTPPOST(CommandName.MeetingResponse, null, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(meetingResponseResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call Provision command.
        /// </summary>
        /// <param name="policyKey">The policy key.</param>
        /// <returns>The response of Provision command.</returns>
        protected SendStringResponse CallProvisionCommand(string policyKey)
        {
            // Create Provision command request.
            ProvisionRequest provisionRequest = Common.CreateProvisionRequest(null, new Request.ProvisionPolicies(), null);
            Request.ProvisionPoliciesPolicy policy = new Request.ProvisionPoliciesPolicy
            {
                PolicyType = "MS-EAS-Provisioning-WBXML"
            };

            if (!string.IsNullOrEmpty(policyKey))
            {
                policy.PolicyKey = policyKey;
                policy.Status = "1";
            }
            else if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "14.1" || 
                Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "16.0" ||
                Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site) == "16.1")
            {
                // Configure the DeviceInformation.
                Request.DeviceInformation deviceInfomation = new Request.DeviceInformation
                {
                    Set = new Request.DeviceInformationSet
                    {
                        Model = "ASHTTPTest"
                    }
                };
                provisionRequest.RequestData.DeviceInformation = deviceInfomation;
            }

            provisionRequest.RequestData.Policies.Policy = policy;

            // Call Provision command by HTTP POST.
            SendStringResponse provisionResponse = this.HTTPAdapter.HTTPPOST(CommandName.Provision, null, provisionRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(provisionResponse.ResponseDataXML);

            return provisionResponse;
        }

        /// <summary>
        /// Call ItemOperations command.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder.</param>
        /// <param name="itemServerId">The ServerId of the email to be fetched.</param>
        /// <param name="acceptMultiPart">Whether need the server to return the requested item in multipart format. It is used in Base64 query value type.</param>
        /// <returns>The response of ItemOperations command.</returns>
        protected SendStringResponse CallItemOperationsCommand(string collectionId, string itemServerId, bool acceptMultiPart)
        {
            Request.ItemOperationsFetch fetch = new Request.ItemOperationsFetch
            {
                Store = "Mailbox",
                CollectionId = collectionId,
                ServerId = itemServerId
            };

            ItemOperationsRequest itemOperationsRequest = Common.CreateItemOperationsRequest(new object[] { fetch });
            Dictionary<CmdParameterName, object> commandParameter = null;
            if (acceptMultiPart)
            {
                commandParameter = new Dictionary<CmdParameterName, object> { { CmdParameterName.Options, 2 } };
            }

            // Call ItemOperations command by HTTP POST.
            SendStringResponse itemOperationsResponse = this.HTTPAdapter.HTTPPOST(CommandName.ItemOperations, commandParameter, itemOperationsRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(itemOperationsResponse.ResponseDataXML);

            return itemOperationsResponse;
        }

        /// <summary>
        /// Call FolderCreate command to create a sub folder under the specific parent folder.
        /// </summary>
        /// <param name="syncKey">The SyncKey of the latest sync.</param>
        /// <param name="createdFolder">The folder name of the sub folder to create.</param>
        /// <param name="parentFolder">The ServerId of the parent folder.</param>
        /// <returns>The response of FolderCreate command.</returns>
        protected FolderCreateResponse CallFolderCreateCommand(string syncKey, string createdFolder, string parentFolder)
        {
            FolderCreateRequest folderCreateRequest = Common.CreateFolderCreateRequest(syncKey, (byte)12, createdFolder, parentFolder);
            SendStringResponse folderCreateResponseString = this.HTTPAdapter.HTTPPOST(CommandName.FolderCreate, null, folderCreateRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(folderCreateResponseString.ResponseDataXML);

            FolderCreateResponse folderCreateResponse = new FolderCreateResponse
            {
                ResponseDataXML = folderCreateResponseString.ResponseDataXML
            };
            folderCreateResponse.DeserializeResponseData();

            return folderCreateResponse;
        }

        /// <summary>
        /// Call FolderUpdate command to rename a folder and move it to another location.
        /// </summary>
        /// <param name="syncKey">The SyncKey of the latest sync.</param>
        /// <param name="itemServerId">The ServerId of the folder to update.</param>
        /// <param name="folderName">The new name of the folder.</param>
        /// <param name="parentId">The ServerId of the parent folder.</param>
        protected void CallFolderUpdateCommand(string syncKey, string itemServerId, string folderName, string parentId)
        {
            FolderUpdateRequest folderUpdateRequest = Common.CreateFolderUpdateRequest(syncKey, itemServerId, folderName, parentId);
            SendStringResponse folderUpdateResponse = this.HTTPAdapter.HTTPPOST(CommandName.FolderUpdate, null, folderUpdateRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(folderUpdateResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call FolderDelete command to delete a folder from server.
        /// </summary>
        /// <param name="syncKey">The SyncKey of the latest sync.</param>
        /// <param name="itemServerId">The ServerId of the folder to delete.</param>
        protected void CallFolderDeleteCommand(string syncKey, string itemServerId)
        {
            FolderDeleteRequest folderDeleteRequest = Common.CreateFolderDeleteRequest(syncKey, itemServerId);
            SendStringResponse folderDeleteResponse = this.HTTPAdapter.HTTPPOST(CommandName.FolderDelete, null, folderDeleteRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(folderDeleteResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call MoveItems command to move an item to another location.
        /// </summary>
        /// <param name="itemServerId">The ServerId of the email to move.</param>
        /// <param name="sourceFolder">The ServerId of the folder to move from.</param>
        /// <param name="destinationFolder">The ServerId of the folder to move to.</param>
        protected void CallMoveItemsCommand(string itemServerId, string sourceFolder, string destinationFolder)
        {
            Request.MoveItemsMove[] moveItemsMove = new Request.MoveItemsMove[1];
            moveItemsMove[0] = new Request.MoveItemsMove
            {
                DstFldId = destinationFolder,
                SrcFldId = sourceFolder,
                SrcMsgId = itemServerId
            };

            MoveItemsRequest request = Common.CreateMoveItemsRequest(moveItemsMove);
            SendStringResponse moveItemsResponse = this.HTTPAdapter.HTTPPOST(CommandName.MoveItems, null, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(moveItemsResponse.ResponseDataXML);
        }

        /// <summary>
        /// Call Ping command to request that the server monitor specified folder for changes.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the server monitor specified folder.</param>
        /// <returns>The response of Ping command.</returns>
        protected PingResponse CallPingCommand(string collectionId)
        {
            // Create the Ping command request
            Request.PingFolder[] pingFolders = new Request.PingFolder[1];
            pingFolders[0] = new Request.PingFolder { Class = Request.PingFolderClass.Email, Id = collectionId };
            PingRequest pingRequest = Common.CreatePingRequest(pingFolders, "80");

            // Call Ping command by HTTP POST.
            SendStringResponse pingResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Ping, null, pingRequest.GetRequestDataSerializedXML());

            PingResponse pingResponse = new PingResponse { ResponseDataXML = pingResponseString.ResponseDataXML };
            pingResponse.DeserializeResponseData();

            return pingResponse;
        }

        /// <summary>
        /// Verify the GetAttachment command parameter.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        /// <returns>The response of GetAttachment command.</returns>
        protected SendStringResponse VerifyGetAttachmentsCommandParameter(QueryValueType queryValueType)
        {
            #region Call SendMail command to send an email with attachment.
            // Call ConfigureRequestPrefixFields to change the QueryValueType.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, queryValueType.ToString());
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call FolderSync command to synchronize the collection hierarchy.
            this.CallFolderSyncCommand();

            // Create SendMail command request.
            string sendMailSubject = Common.GenerateResourceName(Site, "EmailWithAttachment");
            string mimeFileName = "EmailWithAttachment.txt";
            Site.Assert.IsTrue(File.Exists(mimeFileName), "The mime file {0} with attachment content should exist.", mimeFileName);
            string mime = "To: " + this.UserTwoInformation.UserName + "\n" + "Subject: " + sendMailSubject + "\n" + File.ReadAllText(mimeFileName);
            SendMailRequest request = Common.CreateSendMailRequest(null, Guid.NewGuid().ToString("N"), mime);

            // Call SendMail command by HTTP POST.
            SendStringResponse sendMailResponse = this.HTTPAdapter.HTTPPOST(CommandName.SendMail, null, request.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(sendMailResponse.ResponseDataXML);
            #endregion

            #region Call GetAttachment command to get the attachment from received email.
            // Switch the user to User2 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserTwoInformation, true);

            // Add the ServerId of the received email to CreatedItemsForUserTwo collection to delete.
            this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, sendMailSubject);

            // Loop to call Sync command to get the FileReference of the attachment.
            Response.SearchResponseStoreResult searchResult = this.LoopToSearchItem(this.UserTwoInformation.InboxCollectionId, sendMailSubject);
            string fileReference = this.GetFileReferenceFromSearchResult(searchResult);
            Site.Assert.IsNotNull(fileReference, "The FileReference of the attachment in email {0} should be got.", sendMailSubject);

            // Call GetAttachment command using the FileReference as AttachmentName command parameter.
            Dictionary<CmdParameterName, object> commandParameters = new Dictionary<CmdParameterName, object>
            {
                {
                    CmdParameterName.AttachmentName, fileReference
                }
            };

            SendStringResponse getAttachmentResponse = this.HTTPAdapter.HTTPPOST(CommandName.GetAttachment, commandParameters, Common.CreateGetAttachmentRequest().GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(getAttachmentResponse.ResponseDataXML);
            #endregion

            #region Reset the request prefix.
            // Call ConfigureRequestPrefixFields to reset the QueryValueType and credential.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            this.SwitchUser(this.UserOneInformation, false);
            #endregion

            return getAttachmentResponse;
        }

        /// <summary>
        /// Verify the SaveInSent command parameter.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        /// <param name="sendMailSaveInSent">The value of the SaveInSent command parameter in SendMail command.</param>
        /// <param name="smartForwardSaveInSent">The value of the SaveInSent command parameter in SmartForward command.</param>
        /// <param name="smartReplySaveInSent">The value of the SaveInSent command parameter in SmartReply command.</param>
        protected void VerifySaveInSentCommandParameter(QueryValueType queryValueType, string sendMailSaveInSent, string smartForwardSaveInSent, string smartReplySaveInSent)
        {
            #region Call FolderSync and Sync commands to get the initial status.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            string sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
            string smartForwardSubject = Common.GenerateResourceName(Site, "SmartForward");
            string smartReplySubject = Common.GenerateResourceName(Site, "SmartReply");

            // Call ConfigureRequestPrefixFields to change the QueryValueType.
            requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, queryValueType.ToString());
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call FolderSync command to synchronize the collection hierarchy.
            FolderSyncResponse folderSyncResponse = this.CallFolderSyncCommand();

            // Call Sync command to synchronize changes in Sent Items folder.
            SyncResponse syncResponse = this.CallInitialSyncCommand(Common.GetDefaultFolderServerId(folderSyncResponse, FolderType.SentItems, Site));

            if (Common.IsRequirementEnabled(1202, this.Site))
            {
                bool isR1202Satisfied = TestSuiteHelper.VerifySyncRequiredResponseHeaders(folderSyncResponse.Headers.AllKeys);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R1202");
                Site.Log.Add(LogEntryKind.Debug, "The FolderSync command response header is {0}." + folderSyncResponse.Headers.ToString());
                Site.Log.Add(LogEntryKind.Debug, "The Sync command response header is {0}." + syncResponse.Headers.ToString());

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R1202
                // If the FolderSync response header contains X-MS-RP, MS-ASProtocolCommands, and MS-ASProtocolVersions headers, this requirement can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isR1202Satisfied,
                    1202,
                    @"[In Appendix A: Product Behavior] If the client sends a request to synchronize the folder hierarchy with a synchronization key of 0, the implementation does include an X-MS-RP header, MS-ASProtocolCommands header, and a MS-ASProtocolVersions header in its response to the client. (Exchange 2007 SP1 and above follow this behavior.)");
            }
            #endregion

            #region Call SendMail command with setting SaveInSent command parameter.
            string userOneMailboxAddress = Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain);
            string userTwoMailboxAddress = Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain);

            // Call SendMail command to send email to User2 with setting SaveInSent command parameter.
            this.CallSendMailCommand(userOneMailboxAddress, userTwoMailboxAddress, sendMailSubject, sendMailSaveInSent);

            // Add the received email to CreatedItemsForUserTwo collection to delete.
            this.SwitchUser(this.UserTwoInformation, true);
            this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, sendMailSubject);
            this.SwitchUser(this.UserOneInformation, true);

            // Call Sync command to get the ServerId of the saved email.
            string itemServerId = this.LoopToSyncItem(this.UserOneInformation.SentItemsCollectionId, sendMailSubject, TestSuiteHelper.IsRetryNeeded(sendMailSaveInSent));
            this.VerifySaveInSentCommandParameter_SendMail(queryValueType, itemServerId);

            // Add the saved email to CreatedItemsForUserOne collection to delete.
            if (sendMailSaveInSent == "T" || sendMailSaveInSent == "1")
            {
                this.AddCreatedItemToCollection("User1", this.UserOneInformation.SentItemsCollectionId, sendMailSubject);
            }
            #endregion

            #region Call SmartForward to forward the received email with setting SaveInSent, CollectionId and ItemId command parameters.
            // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserTwoInformation, false);

            // Call Sync command to get the ServerId of the received email.
            itemServerId = this.LoopToSyncItem(this.UserTwoInformation.InboxCollectionId, sendMailSubject, true);

            string userThreeMailboxAddress = Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain);

            // Call SmartForward command to forward the received email to User3 with SaveInSent command parameter.
            this.CallSmartForwardCommand(userTwoMailboxAddress, userThreeMailboxAddress, itemServerId, smartForwardSubject, smartForwardSaveInSent, null, null);

            // Add the received email to CreatedItemsForUserThree collection to delete.
            this.SwitchUser(this.UserThreeInformation, true);
            this.AddCreatedItemToCollection("User3", this.UserThreeInformation.InboxCollectionId, smartForwardSubject);
            this.SwitchUser(this.UserTwoInformation, false);

            // Call Sync command to get the ServerId of the saved email.
            itemServerId = this.LoopToSyncItem(this.UserTwoInformation.SentItemsCollectionId, smartForwardSubject, TestSuiteHelper.IsRetryNeeded(smartForwardSaveInSent));
            this.VerifySaveInSentCommandParameter_SmartForward(queryValueType, itemServerId);

            // Add the saved email to CreatedItemsForUserTwo collection to delete.
            if (smartForwardSaveInSent == "T" || smartForwardSaveInSent == "1")
            {
                this.AddCreatedItemToCollection("User2", this.UserTwoInformation.SentItemsCollectionId, smartForwardSubject);
            }
            #endregion

            #region Call SmartReply to reply the received email with setting SaveInSent, CollectionId and ItemId command parameters.
            // Call ConfigureRequestPrefixFields to switch the credential to User3 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserThreeInformation, false);

            // Call Sync command to get the ServerId of the received email.
            itemServerId = this.LoopToSyncItem(this.UserThreeInformation.InboxCollectionId, smartForwardSubject, true);
            if (queryValueType == QueryValueType.PlainText)
            {
                this.VerifyCollectionIdItemIdCommandParameters_SmartForward(itemServerId);
            }
    
            // Call SmartReply command to reply to the received email to User2 with SaveInSent command parameter.
            this.CallSmartReplyCommand(userThreeMailboxAddress, userTwoMailboxAddress, itemServerId, smartReplySubject, smartReplySaveInSent, null, null);

            // Add the received email to CreatedItemsForUserTwo collection to delete.
            this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, smartReplySubject);

            // Call Sync command to get the ServerId of the saved email.
            itemServerId = this.LoopToSyncItem(this.UserThreeInformation.SentItemsCollectionId, smartReplySubject, TestSuiteHelper.IsRetryNeeded(smartReplySaveInSent));
            this.VerifySaveInSentCommandParameter_SmartReply(queryValueType, itemServerId);

            // Add the saved email to CreatedItemsForUserThree collection to delete.
            if (smartReplySaveInSent == "T" || smartReplySaveInSent == "1")
            {
                this.AddCreatedItemToCollection("User3", this.UserThreeInformation.SentItemsCollectionId, smartReplySubject);
            }
            #endregion

            #region Get the received email by SmartReply command and reset the request prefix.
            // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserTwoInformation, false);

            // Call Sync command to get the ServerId of the received email.
            itemServerId = this.LoopToSyncItem(this.UserTwoInformation.InboxCollectionId, smartReplySubject, true);
            if (queryValueType == QueryValueType.PlainText)
            {
                this.VerifyCollectionIdItemIdCommandParameters_SmartReply(itemServerId);
            }

            // Call ConfigureRequestPrefixFields to reset the QueryValueType and credential.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            this.SwitchUser(this.UserOneInformation, false);
            #endregion
        }

        /// <summary>
        /// Verify the LongId command parameter.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        protected void VerifyLongIdCommandParameter(QueryValueType queryValueType)
        {
            #region Call SendMail command to send email to User2.
            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();
            string sendMailSubject = Common.GenerateResourceName(Site, "SendMail");
            string smartForwardSubject = Common.GenerateResourceName(Site, "SmartForward");
            string smartReplySubject = Common.GenerateResourceName(Site, "SmartReply");

            // Call ConfigureRequestPrefixFields to change the QueryValueType.
            requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, queryValueType.ToString());
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);

            // Call FolderSync command to synchronize the collection hierarchy.
            this.CallFolderSyncCommand();
            string userOneMailboxAddress = Common.GetMailAddress(this.UserOneInformation.UserName, this.UserOneInformation.UserDomain);
            string userTwoMailboxAddress = Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain);

            // Call SendMail command to send email to User2.
            this.CallSendMailCommand(userOneMailboxAddress, userTwoMailboxAddress, sendMailSubject, null);
            #endregion

            #region Call SmartForward command with the LongId command parameter returned from Search response.
            // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserTwoInformation, true);

            // Add the received email to the created items collection of User2.
            this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, sendMailSubject);

            // Call Search command to find the received email and get the LongId of it.
            Response.SearchResponseStoreResult searchResult = this.LoopToSearchItem(this.UserTwoInformation.InboxCollectionId, sendMailSubject);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R484");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R484
            // SendMail command executed successfully and the received email can be got from server, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                searchResult,
                484,
                @"[In Command Codes] [Command] SendMail sends mail to the server.");

            string userThreeMailboxAddress = Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain);

            // Call SmartForward command to forward the received email to User3 with the LongId command parameter returned from Search response.
            this.CallSmartForwardCommand(userTwoMailboxAddress, userThreeMailboxAddress, null, smartForwardSubject, null, searchResult.LongId, null);
            #endregion

            #region Get the received email.
            // Call ConfigureRequestPrefixFields to switch the credential to User3 and synchronize the collection hierarchy.
            this.SwitchUser(this.UserThreeInformation, true);

            // Add the received email to CreatedItemsForUserThree collection to delete.
            this.AddCreatedItemToCollection("User3", this.UserThreeInformation.InboxCollectionId, smartForwardSubject);

            // Call Sync command to get the ServerId of the received email.
            string itemServerId = this.LoopToSyncItem(this.UserThreeInformation.InboxCollectionId, smartForwardSubject, true);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R488");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R488
            // SmartForward command executed successfully with the ServerId of the original email, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemServerId,
                488,
                @"[In Command Codes] [Command] SmartForward forwards a Message object without retrieving the full Message object from the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R118");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R118
            // SmartForward command executed successfully with setting LongId command parameter, so this requirement can be captured.
            Site.CaptureRequirement(
                118,
                @"[In Command-Specific URI Parameters] [Parameter] LongId [is used by] SmartForward.");
            #endregion

            #region Call SmartReply command with the LongId command parameter returned from Search response.
            // If the query value type is PlainText, call SmartReply command and verify related requirements.
            if (queryValueType == QueryValueType.PlainText)
            {
                // Call Search command to find the received email and get the LongId of it.
                searchResult = this.LoopToSearchItem(this.UserThreeInformation.InboxCollectionId, smartForwardSubject);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R512");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R512
                // The LongId returned from Search command response is not null, this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    searchResult.LongId,
                    512,
                    @"[In Command-Specific URI Parameters] [Parameter] LongId [is described as] A string that references a result set that was returned in the Search command response.");
              
                // Call SmartReply command to reply the received email to User2 with the LongId command parameter.
                this.CallSmartReplyCommand(userThreeMailboxAddress, userTwoMailboxAddress, null, smartReplySubject, null, searchResult.LongId, null);

                // Call ConfigureRequestPrefixFields to switch the credential to User2 and synchronize the collection hierarchy.
                this.SwitchUser(this.UserTwoInformation, false);

                // Add the received email to CreatedItemsForUserTwo collection to delete.
                this.AddCreatedItemToCollection("User2", this.UserTwoInformation.InboxCollectionId, smartReplySubject);

                // Call Sync command to get the ServerId of the received email.
                itemServerId = this.LoopToSyncItem(this.UserTwoInformation.InboxCollectionId, smartReplySubject, true);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R489");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R489
                // SmartReply command executed successfully with the ServerId of the original email, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    itemServerId,
                    489,
                    @"[In Command Codes] [Command] SmartReply replies to a Message object without retrieving the full Message object from the server.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R528");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R528
                // SmartReply command executed successfully with setting LongId command parameter, so this requirement can be captured.
                Site.CaptureRequirement(
                    528,
                    @"[In Command-Specific URI Parameters] [Parameter] LongId [is used by] SmartReply.");
            }
            #endregion

            #region Reset the query value type and user credential.
            requestPrefix[HTTPPOSTRequestPrefixField.QueryValueType] = Common.GetConfigurationPropertyValue("HeaderEncodingType", this.Site);
            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
            this.SwitchUser(this.UserOneInformation, false);
            #endregion
        }

        /// <summary>
        /// User3 sends meeting request with the specified subject to user2.
        /// </summary>
        /// <param name="meetingRequestSubject">The subject of the meeting request.</param>
        protected void SendMeetingRequest(string meetingRequestSubject)
        {
            // Get the mailboxes.
            string userThreeMailbox = Common.GetMailAddress(this.UserThreeInformation.UserName, this.UserThreeInformation.UserDomain);
            string userTwoMailbox = Common.GetMailAddress(this.UserTwoInformation.UserName, this.UserTwoInformation.UserDomain);

            // Create meeting request mime content and call SendMail command to send the meeting request to User2.
            string icalendarContent = TestSuiteHelper.CreateCalendarContent(userThreeMailbox, userTwoMailbox, meetingRequestSubject, "4");
            string mimeContent = Common.CreateMeetingRequestMime(userThreeMailbox, userTwoMailbox, meetingRequestSubject, "Meeting Request", icalendarContent);
            SendMailRequest sendMailRequest = Common.CreateSendMailRequest(null, System.Guid.NewGuid().ToString(), mimeContent);
            SendStringResponse sendMailResponse = this.HTTPAdapter.HTTPPOST(CommandName.SendMail, null, sendMailRequest.GetRequestDataSerializedXML());

            // Check the command is executed successfully.
            this.CheckResponseStatus(sendMailResponse.ResponseDataXML);
        }

        /// <summary>
        /// Loop to find the specific email by Sync command and return its serverId.
        /// </summary>
        /// <param name="folderCollectionId">The CollectionId of the folder to sync.</param>
        /// <param name="subject">The subject of the email to get.</param>
        /// <param name="isRetryNeeded">Whether retry is needed when get the specified email.</param>
        /// <returns>The ServerId from the email to get.</returns>
        protected string LoopToSyncItem(string folderCollectionId, string subject, bool isRetryNeeded)
        {
            string itemServerId = null;
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
            do
            {
                // Wait for the email received.
                System.Threading.Thread.Sleep(waitTime);

                SyncStore syncResponse = this.CallSyncCommand(folderCollectionId);
                if (syncResponse.SyncKey != null && syncResponse.AddCommands.Count != 0)
                {
                    // Get the ServerId from the response.
                    itemServerId = TestSuiteHelper.GetServerIdFromSyncResponse(syncResponse, subject);
                }

                counter++;
            }
            while (isRetryNeeded && counter < upperBound && itemServerId == null);

            // If the item is supposed to exist, assert the ServerId of it could be got successfully.
            if (isRetryNeeded)
            {
                // Assert the ServerId of the item.
                Site.Assert.IsNotNull(itemServerId, "Can't find the item with subject {0} after retrying {1} times.", subject, counter);
            }

            return itemServerId;
        }

        /// <summary>
        /// Loop to get the LongId of the specific email by Search command.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to search.</param>
        /// <param name="subject">The subject of the email to get.</param>
        /// <returns>The LongId from the email to get.</returns>
        protected Response.SearchResponseStoreResult LoopToSearchItem(string collectionId, string subject)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", Site));
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", Site));
            Response.SearchResponseStoreResult result = null;
            while (counter < upperBound)
            {
                // Wait for the item received.
                System.Threading.Thread.Sleep(waitTime);

                SearchResponse searchResponse = this.CallSearchCommand(collectionId, subject);

                Site.Assert.IsNotNull(searchResponse.ResponseData, "The Search response should not be null.");
                if (!string.Equals(searchResponse.ResponseData.Status, "10", StringComparison.Ordinal))
                {
                    Site.Assert.AreEqual<string>("1", searchResponse.ResponseData.Status, "As a child element of Search, the Status element should be 1 which means success.");
                    Site.Assert.IsNotNull(searchResponse.ResponseData.Response, "The Response element should not be null.");
                    Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store, "The Store element should not be null.");
                    Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Result, "The Result element in Search response should not be null.");

                    Response.SearchResponseStoreResult[] results = searchResponse.ResponseData.Response.Store.Result;
                    if (results.Length > 0 && results[0].Properties != null)
                    {
                        Site.Assert.AreEqual<int>(1, results.Length, "There should be one result in the response, actual {0}.", results.Length);
                        result = results[0];
                        break;
                    }
                }

                counter++;
            }

            // Assert the search result.
            Site.Assert.IsNotNull(result, "Can't find the item with subject {0} after retrying {1} times.", subject, counter);
            return result;
        }

        /// <summary>
        /// Check whether the string returned from response is in xml format.
        /// </summary>
        /// <param name="returnedValue">The string returned from response.</param>
        /// <returns>If the string is in xml format, return true; otherwise, return false.</returns>
        protected bool IsXml(string returnedValue)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(returnedValue);
            }
            catch (XmlException exception)
            {
                Site.Log.Add(LogEntryKind.Debug, string.Format("The string value is not xml and exception {0} is got.", exception.Message));
                return false;
            }

            return true;
        }

        /// <summary>
        /// Check the value of Status in response data xml.
        /// </summary>
        /// <param name="responseXML">The response data xml.</param>
        protected void CheckResponseStatus(string responseXML)
        {
            if (!string.IsNullOrEmpty(responseXML) && this.IsXml(responseXML))
            {
                Regex regexClass = new Regex(@"<Status>(.*?)</Status>", RegexOptions.Singleline);
                Match match = regexClass.Match(responseXML);
                string statusValue = match.Groups[1].Value;

                if (responseXML.Contains("MoveItems"))
                {
                    Site.Assert.AreEqual<string>("3", statusValue, "The value of Status element in response of MoveItems should be 3, actual is {0}.", statusValue);
                }
                else if (responseXML.Contains("Ping"))
                {
                    Site.Assert.AreEqual<string>("2", statusValue, "The value of Status element in response of Ping should be 2, actual is {0}.", statusValue);
                }
                else
                {
                    Site.Assert.AreEqual<string>("1", statusValue, "The value of Status element in response which is not MoveItems or Ping should be 1, actual is {0}.", statusValue);
                }
            }
        }

        /// <summary>
        /// Add the created item in test cases to the items collections used to clean up.
        /// </summary>
        /// <param name="userName">The identity of the user who has the item.</param>
        /// <param name="parentFolder">The CollectionId of the folder which the item is placed.</param>
        /// <param name="itemSubject">The subject of the item to delete.</param>
        protected void AddCreatedItemToCollection(string userName, string parentFolder, string itemSubject)
        {
            CreatedItems createdItems = new CreatedItems { CollectionId = parentFolder };
            createdItems.ItemSubject.Add(itemSubject);

            // Add the created items to the specified user item collection.
            switch (userName)
            {
                case "User1":
                    this.UserOneInformation.UserCreatedItems.Add(createdItems);
                    break;
                case "User2":
                    this.UserTwoInformation.UserCreatedItems.Add(createdItems);
                    break;
                case "User3":
                    this.UserThreeInformation.UserCreatedItems.Add(createdItems);
                    break;
                default:
                    Site.Assert.Fail("The user name is incorrect.");
                    break;
            }
        }

        /// <summary>
        /// Get the DisplayName or ParentId of the specified folder from FolderSync response.
        /// </summary>
        /// <param name="folderSyncResponse">The FolderSync response.</param>
        /// <param name="folderServerId">The ServerId of the folder to get.</param>
        /// <param name="elementName">The element name of the folder to get, it should be DisplayName or ParentId.</param>
        /// <returns>The value of DisplayName or ParentId.</returns>
        protected string GetFolderFromFolderSyncResponse(ActiveSyncResponseBase<Response.FolderSync> folderSyncResponse, string folderServerId, string elementName)
        {
            Site.Assert.IsNotNull(folderSyncResponse, "The FolderSyncResponse should not be null.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData, "The FolderSyncResponse.ResponseData should not be null.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes, "The FolderSyncResponse.ResponseData.Changes should not be null.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData.Changes.Add, "The FolderSyncResponse.ResponseData.Changes.Add should not be null.");

            foreach (Response.FolderSyncChangesAdd add in folderSyncResponse.ResponseData.Changes.Add)
            {
                if (add.ServerId == folderServerId)
                {
                    if (elementName == "DisplayName")
                    {
                        return add.DisplayName;
                    }
                    else if (elementName == "ParentId")
                    {
                        return add.ParentId;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get element value from Sync response
        /// </summary>
        /// <param name="collectionId">The specified collectionId of the folder</param>
        /// <param name="serverId">The specified serverId</param>
        /// <param name="elementType">The element type</param>
        /// <returns>The element value</returns>
        protected object GetElementValueFromSyncResponse(string collectionId, string serverId, Response.ItemsChoiceType8 elementType)
        {
            // Call initial Sync command.
            SyncResponse syncResponse = this.CallInitialSyncCommand(collectionId);

            Site.Assert.IsNotNull(TestSuiteHelper.LoadSyncResponse(syncResponse), "The Sync response should not be null.");

            SyncRequest syncRequest = TestSuiteHelper.GetSyncRequest(collectionId, TestSuiteHelper.LoadSyncResponse(syncResponse).SyncKey);
            syncRequest.RequestData.WindowSize = "512";

            // If there is MoreAvailable tag in the Sync command response, synchronize again to continue getting items from the server.
            SendStringResponse syncResponseString = null;
            do
            {
                // Get the latest SyncKey.
                syncRequest.RequestData.Collections[0].SyncKey = TestSuiteHelper.LoadSyncResponse(syncResponse).SyncKey;

                // Call Sync command by HTTP POST using the SyncKey returned from last sync.
                syncResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Sync, null, syncRequest.GetRequestDataSerializedXML());

                if (!string.IsNullOrEmpty(syncResponseString.ResponseDataXML))
                {
                    // Check the command is executed successfully.
                    this.CheckResponseStatus(syncResponseString.ResponseDataXML);

                    // Convert from SendStringResponse to SyncResponse.
                    syncResponse = TestSuiteHelper.ConvertSyncResponseFromSendString(syncResponseString);
                }
            }
            while (syncResponseString.ResponseDataXML.Contains("<MoreAvailable />"));

            Response.SyncCollections syncCollections = (Response.SyncCollections)syncResponse.ResponseData.Item;
            Response.SyncCollectionsCollectionCommands commands = null;
            for (int index = 0; index < syncCollections.Collection[0].ItemsElementName.Length; index++)
            {
                if (syncCollections.Collection[0].ItemsElementName[index] == Response.ItemsChoiceType10.Commands)
                {
                    commands = (Response.SyncCollectionsCollectionCommands)syncCollections.Collection[0].Items[index];
                    break;
                }
            }

            foreach (Response.SyncCollectionsCollectionCommandsAdd add in commands.Add)
            {
                if (add.ServerId == serverId)
                {
                    for (int itemIndex = 0; itemIndex < add.ApplicationData.ItemsElementName.Length; itemIndex++)
                    {
                        if (add.ApplicationData.ItemsElementName[itemIndex] == elementType)
                        {
                            return add.ApplicationData.Items[itemIndex];
                        }
                    }
                }
            }

            return null;
        }

        #endregion

        #region Test case capture code
        /// <summary>
        /// Verify the capture codes about SaveInSent command parameter using in SendMail command.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        /// <param name="itemServerId">The ServerId of the email received by the receiver.</param>
        protected void VerifySaveInSentCommandParameter_SendMail(QueryValueType queryValueType, string itemServerId)
        {
            if (queryValueType == QueryValueType.Base64)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R92");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R92
                // The sent email could be found in Sent Items folder when the Options command parameter is set to "1", so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    itemServerId,
                    92,
                    @"[In Command Parameters] [When flag] SaveInSent [value is] 0x01,  [ the meaning is ] setting this flag [SaveInSent] to instruct the server to save the Message object in the user's Sent Items folder.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R93");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R93
                // The SendMail command executes successfully when the Options command parameter is set to "1", so this requirement can be captured.
                Site.CaptureRequirement(
                    93,
                    @"[In Command Parameters] [When flag] SaveInSent [value is] 0x01, [it is ] valid for SendMail.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R533");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R533
                // The SendMail command executes successfully when the Options command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    533,
                    @"[In Command Parameters] [Parameter] Options [ is used by] SendMail.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R486");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R486
                // The sent email could be found in Sent Items folder when the Options command parameter is set to "1", so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    itemServerId,
                    486,
                    @"[In Command Parameters] [Parameter] Options [is described as] A single-byte bitmask that specifies command options.");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R531");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R531
                // The SendMail command executes successfully when the SaveInSent command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    531,
                    @"[In Command-Specific URI Parameters] [Parameter] SaveInSent [is used by] SendMail. ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R121");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R121
                // The sent email could be found in Sent Items folder when the SaveInSent command parameter is set to "T", so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    itemServerId,
                    121,
                    @"[In Command-Specific URI Parameters] Set this parameter [SaveInSent] to T to instruct the server to save the Message object in the user's Sent Items folder [; otherwise, set the parameter [SaveInSent] to F]. ");
            }
        }

        /// <summary>
        /// Verify the capture codes about SaveInSent command parameter using in SmartForward command.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        /// <param name="itemServerId">The ServerId of the email received by the receiver.</param>
        protected void VerifySaveInSentCommandParameter_SmartForward(QueryValueType queryValueType, string itemServerId)
        {
            if (queryValueType == QueryValueType.Base64)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R532");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R532
                // The SmartForward command executes successfully when the Options command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    532,
                    @"[In Command Parameters] [Parameter] Options [ is used by] SmartForward.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R524");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R524
                // The SmartForward command executes successfully when the Options command parameter is set to "1", so this requirement can be captured.
                Site.CaptureRequirement(
                    524,
                    @"[In Command Parameters] [When flag] SaveInSent [value is] 0x01, [it is ] valid for SmartForward.");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R437");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R437
                // The ServerId of the saved item is null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNull(
                    itemServerId,
                    437,
                    @"[In Command-Specific URI Parameters] [Set this parameter [SaveInSent] to T to instruct the server to save the Message object in the user's Sent Items folder;] otherwise, set the parameter [SaveInSent] to F.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R120");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R120
                // The SmartForward command executes successfully when the SaveInSent command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    120,
                    @"[In Command-Specific URI Parameters] [Parameter] SaveInSent [is used by] SmartForward.");
            }
        }

        /// <summary>
        /// Verify the capture codes about SaveInSent command parameter using in SmartReply command.
        /// </summary>
        /// <param name="queryValueType">The query value type of the request.</param>
        /// <param name="itemServerId">The ServerId of the email received by the receiver.</param>
        protected void VerifySaveInSentCommandParameter_SmartReply(QueryValueType queryValueType, string itemServerId)
        {
            if (queryValueType == QueryValueType.Base64)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R88");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R88
                // The SmartReply command executes successfully when the Options command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    88,
                    @"[In Command Parameters] [Parameter] Options [ is used by] SmartReply.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R525");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R525
                // The SmartReply command executes successfully when the Options command parameter is set to "1", so this requirement can be captured.
                Site.CaptureRequirement(
                    525,
                    @"[In Command Parameters] [When flag] SaveInSent [value is] 0x01, [it is ] valid for SmartReply.");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R530");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R530
                // The SmartReply command executes successfully when the SaveInSent command parameter is set, so this requirement can be captured.
                Site.CaptureRequirement(
                    530,
                    @"[In Command-Specific URI Parameters] [Parameter] SaveInSent [is used by] SmartReply. ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R122");

                // Verify MS-ASHTTP requirement: MS-ASHTTP_R122
                // The sent email could not be found in Sent Items folder when the SaveInSent command parameter is not set, so this requirement can be captured.
                Site.CaptureRequirementIfIsNull(
                    itemServerId,
                    122,
                    @"[In Command-Specific URI Parameters] The SaveInSent parameter is set to F by default. ");
            }
        }

        /// <summary>
        /// Verify the capture codes about CollectionId and ItemId command parameters using in SmartForward command.
        /// </summary>
        /// <param name="itemServerId">The ServerId of the email received by the receiver.</param>
        protected void VerifyCollectionIdItemIdCommandParameters_SmartForward(string itemServerId)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R517");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R517
            // The forwarded email could be received when the CollectionId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemServerId,
                517,
                @"[In Command-Specific URI Parameters] [Parameter] CollectionId [is described as] A string that specifies the server ID of the folder that contains the Message object to be forwarded [or replied to]. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R511");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R511
            // The forwarded email could be received when the ItemId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemServerId,
                511,
                @"[In Command-Specific URI Parameters] [Parameter] ItemId [is described as] A string that specifies the server ID of the Message object to be forwarded [or replied to]. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R116");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R116
            // The SmartForward command executes successfully when the CollectionId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirement(
                116,
                @"[In Command-Specific URI Parameters] [Parameter] CollectionId [is used by] SmartForward. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R117");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R117
            // The SmartForward command executes successfully when the ItemId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirement(
                117,
                @"[In Command-Specific URI Parameters] [Parameter] ItemId [is used by] SmartForward. ");
        }

        /// <summary>
        /// Verify the capture codes about CollectionId and ItemId command parameters using in SmartReply command.
        /// </summary>
        /// <param name="itemServerId">The ServerId of the email received by the receiver.</param>
        protected void VerifyCollectionIdItemIdCommandParameters_SmartReply(string itemServerId)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R526");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R526
            // The SmartReply command executes successfully when the CollectionId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirement(
                526,
                @"[In Command-Specific URI Parameters] [Parameter] CollectionId [is used by] SmartReply. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R527");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R527
            // The SmartReply command executes successfully when the ItemId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirement(
                527,
                @"[In Command-Specific URI Parameters] [Parameter] ItemId [is used by] SmartReply. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R435");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R435
            // The replied email could be received when the CollectionId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemServerId,
                435,
                @"[In Command-Specific URI Parameters] [Parameter] CollectionId [is described as] A string that specifies the server ID of the folder that contains the Message object to be [forwarded or] replied to. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASHTTP_R122");

            // Verify MS-ASHTTP requirement: MS-ASHTTP_R122
            // The replied email could be received when the ItemId command parameter is set, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                itemServerId,
                436,
                @"[In Command-Specific URI Parameters] [Parameter] ItemId [is described as] A string that specifies the server ID of the Message object to be [forwarded or] replied to. ");
        }
        #endregion

        #region Private method
        /// <summary>
        /// Initialize the HTTP POST request prefix field.
        /// </summary>
        private void InitializeRequestPrefixField()
        {
            string queryValueType = Common.GetConfigurationPropertyValue("HeaderEncodingType", Site);

            IDictionary<HTTPPOSTRequestPrefixField, string> requestPrefix = new Dictionary<HTTPPOSTRequestPrefixField, string>();

            if (string.Equals(queryValueType, "Base64", StringComparison.CurrentCultureIgnoreCase))
            {
                requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.Base64.ToString());
            }
            else if (string.Equals(queryValueType, "PlainText", StringComparison.CurrentCultureIgnoreCase))
            {
                requestPrefix.Add(HTTPPOSTRequestPrefixField.QueryValueType, QueryValueType.PlainText.ToString());
            }

            string activeSyncProtocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", Site);
            requestPrefix.Add(HTTPPOSTRequestPrefixField.ActiveSyncProtocolVersion, Common.ConvertActiveSyncProtocolVersion(activeSyncProtocolVersion, this.Site));

            requestPrefix.Add(HTTPPOSTRequestPrefixField.AcceptEncoding, null);
            requestPrefix.Add(HTTPPOSTRequestPrefixField.PrefixOfURI, Common.GetConfigurationPropertyValue("TransportType", this.Site));
            requestPrefix.Add(HTTPPOSTRequestPrefixField.Host, Common.GetConfigurationPropertyValue("SutComputerName", this.Site));
            requestPrefix.Add(HTTPPOSTRequestPrefixField.AcceptMultiPart, null);
            requestPrefix.Add(HTTPPOSTRequestPrefixField.UserAgent, null);
            requestPrefix.Add(HTTPPOSTRequestPrefixField.PolicyKey, string.Empty);

            this.HTTPAdapter.ConfigureRequestPrefixFields(requestPrefix);
        }

        /// <summary>
        /// Delete the specified item.
        /// </summary>
        /// <param name="itemsToDelete">The collection of the items to delete.</param>
        private void DeleteCreatedItems(Collection<CreatedItems> itemsToDelete)
        {
            foreach (CreatedItems itemToDelete in itemsToDelete)
            {
                SyncStore syncResponse = this.CallSyncCommand(itemToDelete.CollectionId);
                Site.Assert.AreNotEqual<int>(0, syncResponse.AddCommands.Count, "There is not items added in {0} folder.", itemToDelete.CollectionId);
                List<Request.SyncCollectionDelete> deleteList = new List<Request.SyncCollectionDelete>();
                foreach (string itemSubject in itemToDelete.ItemSubject)
                {
                    Request.SyncCollectionDelete appData = new Request.SyncCollectionDelete
                    {
                        ServerId = TestSuiteHelper.GetServerIdFromSyncResponse(syncResponse, itemSubject)
                    };

                    Site.Assert.IsNotNull(appData.ServerId, "The item with subject {0} in {1} folder is not found.", itemSubject, itemToDelete.CollectionId);
                    deleteList.Add(appData);
                }

                // Create the Sync command request.
                Request.SyncCollection[] syncCollections = new Request.SyncCollection[1];
                syncCollections[0] = new Request.SyncCollection
                {
                    CollectionId = itemToDelete.CollectionId,
                    SyncKey = syncResponse.SyncKey,
                    Commands = deleteList.ToArray(),
                    DeletesAsMoves = false,
                    DeletesAsMovesSpecified = true
                };
                SyncRequest syncRequest = Common.CreateSyncRequest(syncCollections);

                // Call Sync command to delete the specified item.
                SendStringResponse syncResponseString = this.HTTPAdapter.HTTPPOST(CommandName.Sync, null, syncRequest.GetRequestDataSerializedXML());

                // Check the command is executed successfully.
                this.CheckResponseStatus(syncResponseString.ResponseDataXML);
            }
        }

        /// <summary>
        /// Get the FileReference element of an attachment from search response.
        /// </summary>
        /// <param name="result">The result of Search command.</param>
        /// <returns>The value of FileReference element.</returns>
        private string GetFileReferenceFromSearchResult(Response.SearchResponseStoreResult result)
        {
            string fileReference = null;
            if (result.Class != null)
            {
                for (int i = 0; i < result.Properties.ItemsElementName.Length; i++)
                {
                    if (result.Properties.ItemsElementName[i] == Response.ItemsChoiceType6.Attachments)
                    {
                        Response.Attachments emailAttachments = result.Properties.Items[i] as Response.Attachments;

                        // Check the attachment exists in the email.
                        Site.Assert.IsNotNull(emailAttachments, "Attachment should be added in the email.");

                        // The attachment number should be 1.
                        Site.Assert.AreEqual<int>(1, emailAttachments.Items.Length, "There should be one attachment in the email, actual {0}.", emailAttachments.Items.Length);
                        fileReference = ((Response.AttachmentsAttachment)emailAttachments.Items[0]).FileReference;
                    }
                }
            }

            return fileReference;
        }
        #endregion
    }
}