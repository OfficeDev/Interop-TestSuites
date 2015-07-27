//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
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
                          @"[In Custom HTTP Headers] The intent of passing X-WOPI-ProofOld header is to allow the WOPI server to validate that the WOPI request originated from the WOPI client that provided the public key in Discovery via ct_wopi-proof-key.");

            // This test suite will trigger the discovery process before send/receive all messages, and all WOPI messages are depended on discovery process. If send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          72,
                          @"[In Message Processing Events and Sequencing Rules] WOPI Discovery involves a single URI that takes no parameters:
                          HTTP://server/hosting/discovery");

            // This test suite will trigger the discovery process before send/receive all messages, and all WOPI messages are depended on discovery process. If send/receive messages successfully, capture this requirement directly.
            this.Site.CaptureRequirement(
                          74,
                          @"[In HTTP://server/hosting/discovery] The data that describes the supported abilities of the WOPI client and how to invoke these abilities through URIs is provided through the following URI:
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
                          ""CloseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""DisableBrowserCachingOfUserContent"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DisablePrint"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DisableTranslation"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""DownloadUrl"":{""type"":""string"",""default"":"""",""optional"":true},
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
                          ""OwnerId"":{""type"":""string"",""optional"":false},
                          ""PresenceProvider""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PresenceUserId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""PrivacyUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""ProtectInClient"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""ReadOnly"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""RestrictedWebViewOnly"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SHA256"":{""type"":""string"",""optional"":false},
                          ""SignoutUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""Size"":{""type"":""int"",""optional"":false},
                          ""SupportsCoauth"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsCobalt"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsFolders"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsLocks"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsScenarioLinks"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsSecureStore"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""SupportsUpdate"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""TenantId""{""type"":""string"",""default"":"""",""optional"":true},
                          ""TermsOfUseUrl"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""TimeZone""{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserCanAttend"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanNotWriteRelative"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanPresent"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserCanWrite"":{""type"":""bool"",""default"":false,""optional"":true},
                          ""UserFriendlyName"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""UserId"":{""type"":""string"",""default"":"""",""optional"":true},
                          ""Version"":{""type"":""string"",""optional"":false}
                          ""WebEditingDisabled"":{""type"":""bool"",""default"":false,""optional"":true},
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
                this.Site.CaptureRequirement(
                              950,
                              @"[In Response Body] SupportsCoauth is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsCobalt"))
            {
                // Check whether "SupportsCobalt" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              775,
                              @"[In Response Body] SupportsCobalt is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsFolders"))
            {
                // Check whether "SupportsFolders" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              777,
                              @"[In Response Body] SupportsFolders is a Boolean value.");
            }

            if (WOPISerializerHelper.CheckContainItem(jsonString, "SupportsLocks"))
            {
                // Check whether "SupportsLocks" is in JSON. If JSON string contain this item,it must follow JSON response format.
                this.Site.CaptureRequirement(
                              779,
                              @"[In Response Body] SupportsLocks is a Boolean value.");
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
        }

        /// <summary>
        /// This method is used to validate Lock response captures.
        /// </summary>
        /// <param name="response">A parameter represents the response from server.</param>
        private void ValidateLockResponse(WOPIHttpResponse response)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-WOPI_R402");

            // Verify MS-WOPI requirement: MS-WOPI_R402
            // "X-WOPI-OldLock" is not null in response means this item is required.
            this.Site.CaptureRequirementIfIsNotNull(
                          response.Headers["X-WOPI-OldLock"],
                          402,
                          @"[In Lock] X-WOPI-OldLock is Required.");
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
                          @"[In HTTP://server/<...>/wopi*/folder/<id>/children] The contents of a folder being accessed by WOPI are identified by the following URI:
                          HTTP://server/<...>/wopi*/folder/<id>/children");

            // All folder children messages are follow this format. If test suite receive a succeed response, capture this requirement.
            this.Site.CaptureRequirement(
                          700,
                          @"[In HTTP://server/<...>/wopi*/folder/<id>/children] The syntax URI parameters are defined by the following ABNF:
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
                             },
                          }");
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
