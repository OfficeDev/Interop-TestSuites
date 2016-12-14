namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to validate the adapter captures.
    /// </summary>
    public partial class MS_WOPIAdapter
    {
        /// <summary>
        /// This method is used to validate all common captures for adapter. These captures are applied for all messages.
        /// </summary>
        private void ValidateCommonMessageCapture()
        {
            switch (this.currentTransport)
            {
                case TransportProtocol.HTTP:
                    {
                        // If send/receive messages successfully by using HTTP transport, capture this requirement directly.
                        this.Site.CaptureRequirement(
                                      3,
                                      @"[In Transport] Messages MUST be transported using HTTP [or HTTPS] using the default ports for these protocols.");
                        break;
                    }

                case TransportProtocol.HTTPS:
                    {
                        // If send/receive messages successfully by using HTTPS transport, capture this requirement directly.
                        this.Site.CaptureRequirement(
                                      4,
                                      @"[In Transport] Messages MUST be transported using [HTTP or] HTTPS using the default ports for these protocols.");
                        break;
                    }

                default:
                    {
                        this.Site.Assert.Fail("The test suite only support HTTP or HTTPS transport.");
                        break;
                    }
            }

            // All requests sent by this test suite contain the "X-WOPI-Proof" header, if send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          31,
                          @"[In Custom HTTP Headers] The intent of passing X-WOPI-Proof header is to allow the WOPI server to validate that the WOPI request originated from the WOPI client that provided the public key in Discovery via ct_wopi-proof-key.");

            // All requests sent by this test suite contain the "X-WOPI-ProofOld" header, if send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          36,
                          @"[In Custom HTTP Headers] The intent of passing X-WOPI-ProofOld header is to allow the WOPI server to validate that the WOPI request originated from the WOPI client that provided the public key in Discovery via ct_proof-proof-key.");

            // This test suite will trigger the discovery process before send/receive all messages, and all WOPI messages are depended on discovery process. If send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          72,
                          @"[In Message Processing Events and Sequencing Rules] WOPI Discovery involves a single URI that takes no parameters:
                          HTTP://server/hosting/discovery");

            // This test suite will trigger the discovery process before send/receive all messages, and all WOPI messages are depended on discovery process. If send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          74,
                          @"[In HTTP://server/hosting/discovery] The data that describes the supported abilities of the WOPI client and how to call these abilities through URIs is provided through the following URI:
                          HTTP://server/hosting/discovery");
        }

        /// <summary>
        /// This method is used to validate all adapter requirements for file level operations. 
        /// </summary>
        private void ValidateFilesCapture()
        {
            // All file level messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          249,
                          @"[In HTTP://server/<...>/wopi*/files/<id>] The file being accessed by WOPI is identified by the following URI:
                          HTTP://server/<...>/wopi*/files/<id>");

            // All file level messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          250,
                          @"[In HTTP://server/<...>/wopi*/files/<id>] The syntax URI parameters are defined by the following Augmented Backus-Naur Form (ABNF):
                          id = STRING");
        }

        /// <summary>
        /// This method is used to validate all about folder requests capture adapter.
        /// </summary>
        private void ValidateFoldersCapture()
        {
            // All folders level request messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          593,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>] The folder being accessed by WOPI is identified by the following URI:
                          HTTP://server/<...>/wopi*/folders/<id>");

            // All folders level request messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          594,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>] The syntax URI parameters are defined by the following ABNF:
                          id = STRING");
        }

        /// <summary>
        /// This method is used to validate all about file content requests adapter capture.
        /// </summary>
        private void ValidateFileContentCapture()
        {
            // All file content level messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          653,
                          @"[In HTTP://server/<...>/wopi*/files/<id>/contents] The content of a file being accessed by WOPI is identified by the following URI:
                          HTTP://server/<...>/wopi*/files/<id>/contents");

            // All file content level messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          654,
                          @"[In HTTP://server/<...>/wopi*/files/<id>/contents] The syntax URI parameters are defined by the following ABNF:
                          id = STRING");
        }

        /// <summary>
        /// This method is used to validate CheckFileInfo response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateCheckFileInfoResponse(WOPIHttpResponse response)
        {
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(response);
            CheckFileInfo checkFileInfo = WOPISerializerHelper.JsonToObject<CheckFileInfo>(jsonString);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-WOPI_R273");

            // Verify MS-WOPI requirement: MS-WOPI_R273
            // If the value is not null indicating the JSON string has been converted to CheckFileInfo type object successfully.
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFileInfo,
                          273,
                          @"[In Response Body] The response body is JavaScript Object Notation (JSON) (as specified in [RFC4627]) with the following parameters:
                          JSON:
                          {
                          ""AllowExternalMarketplace"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""BaseFileName"":{""type"":""string"",""optional"":false},
                          ""BreadcrumbBrandName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbBrandUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbDocName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbDocUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbFolderName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbFolderUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""ClientUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""CloseButtonClosesWindow"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""ClosePostMessage"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""CloseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""DisableBrowserCachingOfUserContent"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DisablePrint"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DisableTranslation"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DownloadUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""EditAndReplyUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""EditModePostMessage"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""EditNotificationPostMessage"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""FileExtension"":{""type"":""string"",""default"":"""",""optional"":true}, 
                          ""FileNameMaxLength"":{""type"":""integer"",""default"":250,""optional"":true}, 
                          ""FileSharingPostMessage"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""FileSharingUrl"":{""type"":""string"",""default"":"""",""optional"":true}, 
                          ""FileUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostAuthenticationId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEditUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEmbeddedEditUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEmbeddedViewUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostNotes"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostRestUrl""{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostViewUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""IrmPolicyDescription"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""IrmPolicyTitle"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""LicenseCheckForEditIsEnabled"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""OwnerId"":{""type"":""string"",""optional"":false},
                          ""PostMessageOrigin""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PresenceProvider""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PresenceUserId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PrivacyUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""ProtectInClient"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""ReadOnly"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""RestrictedWebViewOnly"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SHA256"":{""type"":""string"",""optional"":true},
                          ""SignInUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""SignoutUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""Size"":{""type"":""int"",""optional"":false},
                          ""SupportsCoauth"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsCobalt"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsExtendedLockLength"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsFileCreation"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsFolders"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsGetLock"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsLocks"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsRename"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsScenarioLinks"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsSecureStore"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsUpdate"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsUserInfo"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""TenantId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""TermsOfUseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""TimeZone""{""type"":""string"",""default"":"""",""optional"":true},
                          ""UniqueContentId"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserCanAttend"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanNotWriteRelative"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanPresent"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanRename"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanWrite"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserFriendlyName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserId"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserInfo"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""Version"":{""type"":""string"",""optional"":false},
                          ""WebEditingDisabled"":{""type"":""bool"",""default"":false,""optional"":true}
                          }");

            if (WOPISerializerHelper.CheckContainItem(jsonString, "ReadOnly"))
            {
                // Check whether "ReadOnly" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              925,
                              @"[In Response Body] ReadOnly is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsCoauth"))
            {
                // Check whether "SupportsCoauth" is in JSON. If JSON string contain this item,it must follow JSON response format.
                if (Common.IsRequirementEnabled("MS-WOPI", 950001, this.Site))
                {
                    //string jsonItem = "SupportsCoauth";
                    Boolean isVerified = false;
                    if (jsonString.Contains("\"SupportsCoauth\":false"))
                    {
                        isVerified = true;
                    }
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerified, 
                        950001, 
                        @"[In Response Body] Implementation does return the value false for field SupportsCoauth. <1> Section 3.3.5.1.1.2:  SharePoint Foundation 2013, SharePoint Server 2013 and above return the value false for the SupportsCoauth field.");
                }
               if (Common.IsRequirementEnabled("MS-WOPI", 950002, this.Site))
               {
                   //string jsonItem = "SupportsCoauth";
                    Boolean isVerified = false;
                    if (jsonString.Contains("\"SupportsCoauth\":true"))
                    {
                        isVerified = true;
                    }
                    this.Site.CaptureRequirementIfIsTrue(
                        isVerified,
                        950002,
                        @"[In Response Body] Implementation does return the value true for field SupportsCoauth. (SharePoint Server 2010 follows this behavior).");
                }
            }
            
            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsCobalt"))
            {
                // Check whether "SupportsCobalt" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              775,
                              @"[In Response Body] SupportsCobalt is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsExtendedLockLength"))
            {
                // Check whether "SupportsExtendedLockLength" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              776001001,
                              @"[In Response Body] SupportsExtendedLockLength is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsFileCreation"))
            {
                // Check whether "SupportsFileCreation" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              776002001,
                              @"[In Response Body] SupportsFileCreation is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsFolders"))
            {
                // Check whether "SupportsFolders" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              777,
                              @"[In Response Body] SupportsFolders is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsGetLock"))
            {
                // Check whether "SupportsGetLock" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              778001001,
                              @"[In Response Body] SupportsGetLock is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsLocks"))
            {
                // Check whether "SupportsLocks" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              779,
                              @"[In Response Body] SupportsLocks is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsRename"))
            {
                // Check whether "SupportsRename" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              780001001,
                              @"[In Response Body] SupportsRename is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsSecureStore"))
            {
                // Check whether "SupportsSecureStore" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              923,
                              @"[In Response Body] SupportsSecureStore is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsUpdate"))
            {
                // Check whether "SupportsUpdate" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              783,
                              @"[In Response Body] SupportsUpdate is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsUserInfo"))
            {
                // Check whether "SupportsUserInfo" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              784001001,
                              @"[In Response Body] SupportsUserInfo is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "UniqueContentId"))
            {
                // Check whether "UniqueContentId" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              335001001,
                              @"[In Response Body] UniqueContentId is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "UserCanRename"))
            {
                // Check whether "UserCanRename" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              342001001,
                              @"[In Response Body] UserCanRename is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "UserCanNotWriteRelative"))
            {
                // Check whether "UserCanNotWriteRelative" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              920,
                              @"[In Response Body] UserCanNotWriteRelative is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "UserCanWrite"))
            {
                // Check whether "UserCanWrite" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              928,
                              @"[In Response Body] UserCanWrite is a Boolean value.");
            }

            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i=0; i< response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "FileExtension")
                {
                    //Verify MS-WOPI requirement: MS-WOPI_R292004
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        292004,
                        @"[In Response Body] FileExtension: A string specifying the file extension of the file.");
                    
                    //Verify MS-WOPI requirement: MS-WOPI_R292005
                    this.Site.CaptureRequirementIfIsTrue(
                        response.Headers.AllKeys[i].StartsWith("."),
                        292005,
                        @"[In Response Body] This value [FileExtension] MUST begin with a ""."".");
                }
                if (response.Headers.AllKeys[i] == "FileNameMaxLength")
                {
                    int FileNameMaxLength = 0;
                    Boolean isInt = int.TryParse(response.Headers.AllKeys[i], out FileNameMaxLength);
                    //Verify MS-WOPI requirement: MS-WOPI_R292007
                    this.Site.CaptureRequirementIfIsTrue(
                        isInt,
                        292007,
                        @"[In Response Body] FileNameMaxLength: An integer indicating the maximum length for file names, including the file extension, supported by the WOPI server.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");

            this.ValidateURL(checkFileInfo.CloseUrl, "CloseUrl");
            this.ValidateURL(checkFileInfo.DownloadUrl, "DownloadUrl");
            this.ValidateURL(checkFileInfo.FileSharingUrl, "FileSharingUrl");
            this.ValidateURL(checkFileInfo.HostViewUrl, "HostViewUrl");
        }

        /// <summary>
        /// This method is used to validate PutRelativeFile response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidatePutRelativeFileResponse(WOPIHttpResponse response)
        {
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(response);
            PutRelativeFile putRelativeFile = WOPISerializerHelper.JsonToObject<PutRelativeFile>(jsonString);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-WOPI_R382");

            // Verify MS-WOPI requirement: MS-WOPI_R382
            // The object is not null means JSON string can change to object. JSON to object check all require and optional item.
            this.Site.CaptureRequirementIfIsNotNull(
                          putRelativeFile,
                          382,
                          @"[In Response Body] [Name] The response body is JSON (as specified in [RFC4627]) with the following parameters:
                          JSON:
                          {
                          ""Name"":{""type"":""string"",""optional"":false},
                          ""Url"":{""type"":""string"",""default"":"""",""optional"":false},
                          ""HostViewUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEditUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          }");
            this.ValidateURL(putRelativeFile.HostViewUrl, "HostViewUrl");
            this.ValidateURL(putRelativeFile.HostEditUrl, "HostEditUrl");

            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-ValidRelativeTarget")
                {
                    //Verify MS-WOPI requirement: MS-WOPI_R370005
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        370005,
                        @"[In PutRelativeFile] X-WOPI-ValidRelativeTarget is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    //Verify MS-WOPI requirement: MS-WOPI_R370008
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        370008,
                        @"[In PutRelativeFile] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    //Verify MS-WOPI requirement: MS-WOPI_R370013
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        370013,
                        @"[In PutRelativeFile] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");

        }

        /// <summary>
        /// This method is used to validate Lock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateLockResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R401003
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        401003,
                        @"[In Lock] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R401008
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        401008,
                        @"[In Lock] X-WOPI-LockFailureReason is a string.");
                }

                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");

        }

        /// <summary>
        /// This method is used to validate UnLock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateUnLockResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R422005
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        422005,
                        @"[In Unlock] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R422010
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        422010,
                        @"[In Unlock] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate RefreshLock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateRefreshLockResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count;i++ )
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R439005
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        439005,
                        @"[In RefreshLock] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R439010
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        439010,
                        @"[In RefreshLock] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate PutFile response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidatePutFileResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R685005
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        685005,
                        @"[In PutFile] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R685010
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        685010,
                        @"[In PutFile] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate UnlockAndRelock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateUnlockAndRelockResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R460005
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        460005,
                        @"[In UnlockAndRelock] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R460010
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        460010,
                        @"[In UnlockAndRelock] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                isR42001Verified,
                42001,
                @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                isR45001Verified,
                45001,
                @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate GetLock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateGetLockResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R469011
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        469011,
                        @"[In GetLock] X-WOPI-Lock is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R469016
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        469016,
                        @"[In GetLock]  X-WOPI-LockFailureReason is a string");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate RenameFile response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateRenameFileResponse(WOPIHttpResponse response)
        {
            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-InvalidFileNameError")
                {
                    //Verify MS-WOPI requirement: MS-WOPI_R592018
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof (String),
                        response.Headers.AllKeys[i].GetType(),
                        592018,
                        @"[In RenameFile] X-WOPI-InvalidFileNameError is string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-Lock")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R592022
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        592022,
                        @"[In RenameFile] X-WOPI-LockFailureReason is a string.");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-LockFailureReason")
                {
                    // Verify MS-WOPI requirement: MS-WOPI_R592027
                    this.Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(String),
                        response.Headers.AllKeys[i].GetType(),
                        592027,
                        @"[In RenameFile] X-WOPI-LockFailureReason is a string");
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate ReadSecureStore response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateReadSecureStoreResponse(WOPIHttpResponse response)
        {
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(response);

            // If the JSON string can be converted to the ReadSecureStore succeed, that means it match the JSON schema definition.
            WOPISerializerHelper.JsonToObject<ReadSecureStore>(jsonString);

            // If the JSON string can converted to object. The process of "JSON to object" check all require and optional item.
            this.Site.CaptureRequirement(
                          541,
                          @"[In Response Body] The response body is JSON (as specified in [RFC4627]) with the following parameters:
                          JSON:
                          {
                          ""UserName"":{""type"":""string"",""optional"":false},
                          ""Password"":{""type"":""string"",""default"":"""",""optional"":false},
                          ""IsWindowsCredentials"":{""type"":""bool"",""default"":""false"",""optional"":true},
                          ""IsGroup"":{""type"":""bool"",""default"":""false"",""optional"":true},
                          }");

            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        }

        /// <summary>
        /// This method is used to validate CheckFolderInfo response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateCheckFolderInfoResponse(WOPIHttpResponse response)
        {
            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(response);
            CheckFolderInfo checkFolderInfo = WOPISerializerHelper.JsonToObject<CheckFolderInfo>(jsonString);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-WOPI_R602");

            // Verify MS-WOPI requirement: MS-WOPI_R602
            // The object is not null means JSON string can change to object. JSON to object check all require and optional item.
            this.Site.CaptureRequirementIfIsNotNull(
                          checkFolderInfo,
                          602,
                          @"[In Response Body] The response body is JSON (as specified in [RFC4627]) with the following parameters:
                          JSON:
                          {
                          ""FolderName"":{""type"":""string"",""optional"":false},
                          ""BreadcrumbBrandIconUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbBrandName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbBrandUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbDocName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbDocUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbFolderName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""BreadcrumbFolderUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""ClientUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""CloseButtonClosesWindow"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""CloseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""FileSharingUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostAuthenticationId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEditUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEmbeddedEditUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostEmbeddedViewUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""HostViewUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""OwnerId"":{""type"":""string"",""optional"":false},
                          ""PresenceProvider""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PresenceUserId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PrivacyUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""SignoutUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""SupportsSecureStore"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""TenantId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""TermsOfUseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserCanWrite"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserFriendlyName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserId"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""WebEditingDisabled"":{""type"":""bool"",""default"":false,""optional"":true},
                          }");

            this.ValidateURL(checkFolderInfo.CloseUrl, "CloseUrl");
            this.ValidateURL(checkFolderInfo.FileSharingUrl, "FileSharingUrl");
            this.ValidateURL(checkFolderInfo.HostEmbeddedEditUrl, "HostEmbeddedEditUrl");
            this.ValidateURL(checkFolderInfo.HostEmbeddedViewUrl, "HostEmbeddedViewUrl");
            this.ValidateURL(checkFolderInfo.PrivacyUrl, "PrivacyUrl");
            this.ValidateURL(checkFolderInfo.SignoutUrl, "SignoutUrl");

            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        
        }

        /// <summary>
        /// This method is used to validate EnumerateChildren response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateEnumerateChildrenResponse(WOPIHttpResponse response)
        {
            // All folder children messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          699,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>/children] The contents of a folder being accessed by WOPI are identified by the following URI:
HTTP://server/<...>/wopi*/folders/<id>/children");

            // All folder children messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          700,
                          @"[In HTTP://server/<...>/wopi*/folders/<id>/children] The syntax URI parameters are defined by the following ABNF:
id = STRING");

            string jsonString = WOPIResponseHelper.ReadHTTPResponseBodyToString(response);
            EnumerateChildren enumerateChildren = WOPISerializerHelper.JsonToObject<EnumerateChildren>(jsonString);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-WOPI_R712");

            // Verify MS-WOPI requirement: MS-WOPI_R712
            // If the value is not null indicating the JSON string has been converted to EnumerateChildren type object successfully.
            this.Site.CaptureRequirementIfIsNotNull(
                          enumerateChildren,
                          712,
                          @"[In Response Body] The response body is JSON (as specified in [RFC4627]) with the following parameters:
                          JSON:
                          {
                          ""Children"":
                            [{
                              ""Name"":""<name>"",
                              ""Url"":""<url>"",
                              ""Version"":""<version>""
                             }],
                          }");

            Boolean isR42001Verified = false;
            Boolean isR45001Verified = false;
            for (int i = 0; i < response.Headers.Count; i++)
            {
                if (response.Headers.AllKeys[i] == "X-WOPI-ServerVersion")
                {
                    isR42001Verified = true;
                }
                if (response.Headers.AllKeys[i] == "X-WOPI-MachineName")
                {
                    isR45001Verified = true;
                }
            }
            this.Site.CaptureRequirementIfIsTrue(
                    isR42001Verified,
                    42001,
                    @"[In Custom HTTP Headers] Header X-WOPI-ServerVersion [is a string specifying the version of the WOPI server and] MUST be included with all WOPI responses.");

            this.Site.CaptureRequirementIfIsTrue(
                    isR45001Verified,
                    45001,
                    @"[In Custom HTTP Headers] Header X-WOPI-MachineName [is a string specifying the name of the WOPI server and] MUST be included with all WOPI responses, which MUST NOT be used for anything other than logging.");
        
        }

        /// <summary>
        /// A method used to validate the specified URL value whether is a valid URL format. It the value is not a valid URL format, this method will raise an "Assert.Fail" exception.
        /// </summary>
        /// <param name="urlValue">A parameter represents the URL value which will be validated. If this parameter is null or empty value, the method will skip the validation.</param>
        /// <param name="urlName">A parameter represents the URL name of the value which is specified in "urlValue" parameter.</param>
        private void ValidateURL(string urlValue, string urlName)
        {
            if (!string.IsNullOrEmpty(urlValue))
            {
                Uri uriTemp;
                if (!Uri.TryCreate(urlValue, UriKind.Absolute, out uriTemp))
                {
                    this.Site.Assert.Fail(
                        "The {0} should be a valid URL format. Current value is [{1}].",
                        string.IsNullOrEmpty(urlName) ? "URL" : urlName + " value",
                        urlValue);
                }
            }
        }
    }
}