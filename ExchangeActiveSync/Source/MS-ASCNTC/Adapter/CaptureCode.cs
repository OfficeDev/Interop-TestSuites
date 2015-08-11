namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASCNTC.
    /// </summary>
    public partial class MS_ASCNTCAdapter
    {
        #region Verfiy message syntax
        /// <summary>
        /// This method is used to verify the message syntax related requirement.
        /// </summary>
        private void VerifyMessageSyntax()
        {
            // If the validation is successful, then MS-ASCNTC_R4 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R4");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R4
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                4,
                @"[In Message Syntax] The markup that is used by this protocol MUST be well-formed XML, as specified in [XML].");
        }
        #endregion

        #region Verify abstract data model
        /// <summary>
        /// This method is used to verify abstract data model related requirements.
        /// </summary>
        private void VerifyAbstractDataModel()
        {
            // If the validation is successful, then MS-ASCNTC_R436 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R436");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R436
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                436,
                @"[In Abstract Data Model] It[Contact class] is returned by the server to the client as part of a full XML response to the client command requests specified in section 3.1.5.");

            // If the validation is successful, then MS-ASCNTC_R470 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R470");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R470
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                470,
                @"[In Abstract Data Model] It[Contact class] is returned by the server as part of a full XML response to the client requests specified in section 3.1.5.");

            // If the validation is successful, then MS-ASCNTC_R471 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R471");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R471
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                471,
                @"[In Abstract Data Model] Command response: A WBXML formatted message that adheres to the command schemas specified in [MS-ASCMD].");
        }
        #endregion

        #region Verify transport
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R2");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R2
            Site.CaptureRequirement(
                2,
                @"[In Transport] The XML markup that constitutes the request body or the response body that is transmitted between the client and the server uses Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }
        #endregion

        #region Verify ActiveSync command response
        /// <summary>
        /// This method is used to verify the Sync command response related requirements.
        /// </summary>
        /// <param name="syncStore">The Sync result returned from the server.</param>
        private void VerifySyncResponse(SyncStore syncStore)
        {
            if (syncStore.AddElements.Count != 0 || syncStore.ChangeElements.Count != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R473");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R473
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    473,
                    @"[In Synchronizing Contact Data Between Client and Server] [If a client sends a Sync command request to the server]The server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R492");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R492
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    492,
                    @"[In Sync Command Response] When a client uses the Sync command request ([MS-ASCMD] section 2.2.2.19) to synchronize its Contact class items for a specified user with the contacts currently stored by the server, as specified in section 3.1.5.3, the server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R514");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R514
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    514,
                    @"[In Refreshing the Recipient Information Cache] [If a client sends a Sync command request to the server for refreshing]The server responds with a Sync command response ([MS-ASCMD] section 2.2.2.19) that includes only the following elements from the Contact class:
 Email1Address (section 2.2.2.25)
 FileAs (section 2.2.2.28)
 Alias (section 2.2.2.2)
 WeightedRank (section 2.2.2.62)");
            }

            if (syncStore.AddElements.Count != 0)
            {
                foreach (Sync item in syncStore.AddElements)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R493");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R493
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        493,
                        @"[In Sync Command Response] Any of the elements that belong to the Contact class, as specified in section 2.2.2, can be included in a Sync command response as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within either an airsync:Add element ([MS-ASCMD] section 2.2.3.7.2).");

                    this.VerifyContactClassElements(item.Contact);
                }
            }

            if (syncStore.ChangeElements.Count != 0)
            {
                foreach (Sync item in syncStore.ChangeElements)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R494");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R494
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        494,
                        @"[In Sync Command Response] Any of the elements that belong to the Contact class, as specified in section 2.2.2, can be included in a Sync command response as child elements of the airsync:ApplicationData element ([MS-ASCMD] section 2.2.3.11) within [or] an airsync:Change element ([MS-ASCMD] section 2.2.3.24).");

                    this.VerifyContactClassElements(item.Contact);
                }
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }

        /// <summary>
        /// This method is used to verify the ItemOperations command response related requirements.
        /// </summary>
        /// <param name="itemOperationsStore">The ItemOperations result returned from the server.</param>
        private void VerifyItemOperationsResponse(ItemOperationsStore itemOperationsStore)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R477");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R477
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                477,
                @"[In Retrieving Details for Specific Contacts] [If a client sends an ItemOperations command request to the server]The server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R483");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R483
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                483,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.2.8) to retrieve data from the server for one or more contact items, as specified in section 3.1.5.1, the server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");

            // If the validation is successful, then MS-ASCNTC_R486 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R486");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R486
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                486,
                @"[In ItemOperations Command Response] Contact class elements are returned as child elements of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.128) in the ItemOperations command response.");

            foreach (ItemOperations itemOperations in itemOperationsStore.Items)
            {
                if (itemOperations.Contact != null)
                {
                    this.VerifyContactClassElements(itemOperations.Contact);
                }
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }

        /// <summary>
        /// This method is used to verify the Search command response related requirements.
        /// </summary>
        /// <param name="searchStore">The Search result from the server.</param>
        private void VerifySearchResponse(SearchStore searchStore)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R475");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R475
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                475,
                @"[In Searching for Contact Data] [If a client sends a Search command request to the server]The server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R488");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R488
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                488,
                @"[In Search Command Response] When a client uses the Search command request ([MS-ASCMD] section 2.2.2.14) to retrieve Contact class items that match the criteria specified by the client, as specified in section 3.1.5.2, the server responds with a Search command response ([MS-ASCMD] section 2.2.2.14).");

            // If the validation is successful, then MS-ASCNTC_R490 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R490");

            // Verify MS-ASCNTC requirement: MS-ASCNTC_R490
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                490,
                @"[In Search Command Response] Contact class elements are returned as child elements of the search:Properties element ([MS-ASCMD] section 2.2.3.128) in the Search command response.");

            foreach (Search search in searchStore.Results)
            {
                if (search.Contact != null)
                {
                    this.VerifyContactClassElements(search.Contact);
                }
            }

            this.VerifyMessageSyntax();
            this.VerifyAbstractDataModel();
        }
        #endregion

        #region Verify Contact class elements
        /// <summary>
        /// This method is used to verify the contact item related requirements
        /// </summary>
        /// <param name="contact">The contact item synchronized/retrieved/searched from server.</param>
        private void VerifyContactClassElements(Contact contact)
        {
            if (contact.AccountName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R99 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R99");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R99
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    99,
                    @"[In AccountName] The value of this element[contacts2:AccountName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.Anniversary != null)
            {
                // If the validation is successful, then MS-ASCNTC_R109 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R109");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R109
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    109,
                    @"[In Anniversary] The value of this element[Anniversary] is a datetime data type in Coordinated Universal Time (UTC) format, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTime(contact.Anniversary);
            }

            if (contact.AssistantName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R114 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R114");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R114
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    114,
                    @"[In AssistantName] The value of this element[AssistantName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.AssistantPhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R119 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R119");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R119
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    119,
                    @"[In AssistantPhoneNumber] The value of this element[AssistantPhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Birthday != null)
            {
                // If the validation is successful, then MS-ASCNTC_R124 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R124");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R124
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    124,
                    @"[In Birthday] The value of this element[Birthday] is a datetime data type in Coordinated Universal Time (UTC) format, as specified in [MS-ASDTYPE] section 2.3.");

                this.VerifyDateTime(contact.Birthday);
            }

            if (contact.Body != null)
            {
                // If the validation is successful, then MS-ASCNTC_R127 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R127");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R127
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    127,
                    @"[In Body] The airsyncbase:Body element is a container ([MS-ASDTYPE] section 2.2) element.");

                this.VerifyContainer();
            }

            if (contact.BusinessAddressCity != null)
            {
                // If the validation is successful, then MS-ASCNTC_R132 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R132");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R132
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    132,
                    @"[In BusinessAddressCity] The value of this element[BusinessAddressCity] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.BusinessAddressCountry != null)
            {
                // If the validation is successful, then MS-ASCNTC_R137 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R137");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R137
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    137,
                    @"[In BusinessAddressCountry] The value of this element[BusinessAddressCountry] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.BusinessAddressPostalCode != null)
            {
                // If the validation is successful, then MS-ASCNTC_R142 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R142");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R142
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    142,
                    @"[In BusinessAddressPostalCode] The value of this element[BusinessAddressPostalCode] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.BusinessAddressState != null)
            {
                // If the validation is successful, then MS-ASCNTC_R145 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R145");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R145
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    145,
                    @"[In BusinessAddressState] The value of this element[BusinessAddressState] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.BusinessAddressStreet != null)
            {
                // If the validation is successful, then MS-ASCNTC_R150 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R150");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R150
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    150,
                    @"[In BusinessAddressStreet] The value of this element[BusinessAddressStreet] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.BusinessFaxNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R155 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R155");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R155
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    155,
                    @"[In BusinessFaxNumber] The value of this element[BusinessFaxNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.BusinessPhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R160 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R160");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R160
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    160,
                    @"[In BusinessPhoneNumber] The value of this element[BusinessPhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Business2PhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R165 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R165");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R165
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    165,
                    @"[In Business2PhoneNumber] The value of this element[Business2PhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.CarPhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R170 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R170");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R170
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    170,
                    @"[In CarPhoneNumber] The value of this element[CarPhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Categories != null)
            {
                // If the validation is successful, then MS-ASCNTC_R173 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R173");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R173
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    173,
                    @"[In Categories] The Categories element is a container ([MS-ASDTYPE] section 2.2) element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R517");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R517
                Site.CaptureRequirementIfIsNotNull(
                    contact.Categories.Category,
                    517,
                    @"[In Categories] The Categories element has the following child element:Category (section 2.2.2.18): At least one instance of this element is required.");

                this.VerifyContainer();
                this.VerifyCategory(contact.Categories.Category);
            }

            if (contact.Children != null)
            {
                // If the validation is successful, then MS-ASCNTC_R184 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R184");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R184
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    184,
                    @"[In Children] The Children element is a container ([MS-ASDTYPE] section 2.2) element.");

                this.VerifyContainer();
                this.VerifyChild(contact.Children.Child);
            }

            if (contact.CompanyMainPhone != null)
            {
                // If the validation is successful, then MS-ASCNTC_R197 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R197");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R197
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    197,
                    @"[In CompanyMainPhone] The value of this element[contacts2:CompanyMainPhone] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.CompanyName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R202 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R202");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R202
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    202,
                    @"[In CompanyName] The value of this element[CompanyName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.CustomerId != null)
            {
                // If the validation is successful, then MS-ASCNTC_R207 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R207");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R207
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    207,
                    @"[In CustomerId] The value of this element[contacts2:CustomerId] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.Department != null)
            {
                // If the validation is successful, then MS-ASCNTC_R211 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R211");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R211
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    211,
                    @"[In Department] The value of this element[Department] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.Email1Address != null)
            {
                // If the validation is successful, then MS-ASCNTC_R216 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R216");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R216
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    216,
                    @"[In Email1Address] The value of this element[Email1Address] is a string data type, as specified in [MS-ASDTYPE] section 2.6.2.");

                this.VerifyString();
                this.VerifyEmailAddress(contact.Email1Address);
            }

            if (contact.Email2Address != null)
            {
                // If the validation is successful, then MS-ASCNTC_R223 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R223");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R223
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    223,
                    @"[In Email2Address] The value of this element[Email2Address] is a string data type, as specified in [MS-ASDTYPE] section 2.6.2.");

                this.VerifyString();
                this.VerifyEmailAddress(contact.Email2Address);
            }

            if (contact.Email3Address != null)
            {
                // If the validation is successful, then MS-ASCNTC_R228 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R228");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R228
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    228,
                    @"[In Email3Address] The value of this element[Email3Address] is a string data type, as specified in [MS-ASDTYPE] section 2.6.2.");

                this.VerifyString();
                this.VerifyEmailAddress(contact.Email3Address);
            }

            if (contact.FileAs != null)
            {
                // If the validation is successful, then MS-ASCNTC_R233 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R233");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R233
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    233,
                    @"[In FileAs] The value of this element[FileAs ] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.FirstName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R240 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R240");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R240
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    240,
                    @"[In FirstName] The value of this element[FirstName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.GovernmentId != null)
            {
                // If the validation is successful, then MS-ASCNTC_R245 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R245");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R245
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    245,
                    @"[In GovernmentId] The value of this element[contacts2:GovernmentId] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeAddressCity != null)
            {
                // If the validation is successful, then MS-ASCNTC_R250 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R250");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R250
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    250,
                    @"[In HomeAddressCity] The value of this element[HomeAddressCity] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeAddressCountry != null)
            {
                // If the validation is successful, then MS-ASCNTC_R255 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R255");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R255
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    255,
                    @"[In HomeAddressCountry] The value of this element[HomeAddressCountry] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeAddressPostalCode != null)
            {
                // If the validation is successful, then MS-ASCNTC_R260 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R260");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R260
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    260,
                    @"[In HomeAddressPostalCode] The value of this element[HomeAddressPostalCode] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeAddressState != null)
            {
                // If the validation is successful, then MS-ASCNTC_R265 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R265");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R265
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    265,
                    @"[In HomeAddressState] The value of this element[HomeAddressState] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeAddressStreet != null)
            {
                // If the validation is successful, then MS-ASCNTC_R270 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R270");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R270
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    270,
                    @"[In HomeAddressStreet] The value of this element[HomeAddressStreet] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.HomeFaxNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R275 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R275");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R275
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    275,
                    @"[In HomeFaxNumber] The value of this element[HomeFaxNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.HomePhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R280 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R280");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R280
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    280,
                    @"[In HomePhoneNumber] The value of this element[HomePhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Home2PhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R285 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R285");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R285
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    285,
                    @"[In Home2PhoneNumber] The value of this element[Home2PhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.IMAddress != null)
            {
                // If the validation is successful, then MS-ASCNTC_R290 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R290");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R290
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    290,
                    @"[In IMAddress] The value of this element[contacts2:IMAddress] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.IMAddress2 != null)
            {
                // If the validation is successful, then MS-ASCNTC_R295 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R295");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R295
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    295,
                    @"[In IMAddress2] The value of this element[contacts2:IMAddress2] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.IMAddress3 != null)
            {
                // If the validation is successful, then MS-ASCNTC_R300 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R300");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R300
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    300,
                    @"[In IMAddress3] The value of this element[contacts2:IMAddress3] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.JobTitle != null)
            {
                // If the validation is successful, then MS-ASCNTC_R305 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R305");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R305
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    305,
                    @"[In JobTitle] The value of this element[JobTitle] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.LastName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R310 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R310");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R310
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    310,
                    @"[In LastName] The value of this element[LastName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.ManagerName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R315 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R315");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R315
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    315,
                    @"[In ManagerName] The value of this element[contacts2:ManagerName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.MiddleName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R320 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R320");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R320
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    320,
                    @"[In MiddleName] The value of this element[MiddleName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.MMS != null)
            {
                // If the validation is successful, then MS-ASCNTC_R325 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R325");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R325
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    325,
                    @"[In MMS] The value of this element[contacts2:MMS] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.MobilePhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R330 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R330");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R330
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    330,
                    @"[In MobilePhoneNumber] The value of this element[MobilePhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.NickName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R335 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R335");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R335
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    335,
                    @"[In NickName] The value of this element[contacts2:NickName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OfficeLocation != null)
            {
                // If the validation is successful, then MS-ASCNTC_R340 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R340");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R340
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    340,
                    @"[In OfficeLocation] The value of this element[OfficeLocation] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OtherAddressCity != null)
            {
                // If the validation is successful, then MS-ASCNTC_R345 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R345");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R345
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    345,
                    @"[In OtherAddressCity] The value of this element[OtherAddressCity] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OtherAddressCountry != null)
            {
                // If the validation is successful, then MS-ASCNTC_R350 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R350");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R350
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    350,
                    @"[In OtherAddressCountry] The value of this element[OtherAddressCountry] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OtherAddressPostalCode != null)
            {
                // If the validation is successful, then MS-ASCNTC_R355 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R355");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R355
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    355,
                    @"[In OtherAddressPostalCode] The value of this element[OtherAddressPostalCode] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OtherAddressState != null)
            {
                // If the validation is successful, then MS-ASCNTC_R360 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R360");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R360
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    360,
                    @"[In OtherAddressState] The value of this element[OtherAddressState] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.OtherAddressStreet != null)
            {
                // If the validation is successful, then MS-ASCNTC_R365 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R365");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R365
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    365,
                    @"[In OtherAddressStreet] The value of this element[OtherAddressStreet] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.PagerNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R370 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R370");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R370
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    370,
                    @"[In PagerNumber] The value of this element[PagerNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Picture != null)
            {
                if (Common.IsRequirementEnabled(506, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R506");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R506
                    Site.CaptureRequirementIfIsTrue(
                        AdapterHelper.IsPicture(contact.Picture),
                        506,
                        @"[In Appendix B: Product Behavior] The value of the Picture element is a stream that is encoded with base64 encoding. (Exchange Server 2007 and above follow this behavior.)");
                }

                // If the validation is successful, then MS-ASCNTC_R379 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R379");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R379
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    379,
                    @"[In Picture] The value of this element[Picture] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.RadioPhoneNumber != null)
            {
                // If the validation is successful, then MS-ASCNTC_R387 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R387");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R387
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    387,
                    @"[In RadioPhoneNumber] The value of this element[RadioPhoneNumber] is a string data type, as specified in [MS-ASDTYPE] section 2.6.3.");

                this.VerifyString();
            }

            if (contact.Spouse != null)
            {
                // If the validation is successful, then MS-ASCNTC_R392 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R392");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R392
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    392,
                    @"[In Spouse] The value of this element[Spouse] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.Suffix != null)
            {
                // If the validation is successful, then MS-ASCNTC_R397 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R397");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R397
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    397,
                    @"[In Suffix] The value of this element[Suffix] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.Title != null)
            {
                // If the validation is successful, then MS-ASCNTC_R402 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R402");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R402
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    402,
                    @"[In Title] The value of this element[Title] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.WebPage != null)
            {
                // If the validation is successful, then MS-ASCNTC_R407 can be captured.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R407");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R407
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    407,
                    @"[In WebPage] The value of this element[WebPage] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.WeightedRank != null)
            {
                // If the validation is successful, then MS-ASCNTC_R412 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R412");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R412
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    412,
                    @"[In WeightedRank] The value of this element[WeightedRank] is an integer data type, as specified in [MS-ASDTYPE] section 2.5.");

                this.VerifyInteger();
            }

            if (contact.YomiCompanyName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R419 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R419");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R419
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    419,
                    @"[In YomiCompanyName] The value of this element[YomiCompanyName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.YomiFirstName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R424 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R424");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R424
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    424,
                    @"[In YomiFirstName] The value of this element[YomiFirstName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }

            if (contact.YomiLastName != null)
            {
                // If the validation is successful, then MS-ASCNTC_R429 can be captured. 
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R429");

                // Verify MS-ASCNTC requirement: MS-ASCNTC_R429
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    429,
                    @"[In YomiLastName] The value of this element[YomiLastName] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                this.VerifyString();
            }
        }

        /// <summary>
        /// This method is used to verify the Category element.
        /// </summary>
        /// <param name="categoryList">The array of Category element.</param>
        private void VerifyCategory(string[] categoryList)
        {
            if (categoryList != null)
            {
                foreach (string category in categoryList)
                {
                    // If the Category element isn't null, then MS-ASCNTC_R179 can be captured.
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R179");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R179
                    Site.CaptureRequirementIfIsNotNull(
                        category,
                        179,
                        @"[In Category] The Category element is a required child element of the Categories element (section 2.2.2.17) that specifies a category that is assigned to the contact.");

                    // If MS-ASCNTC_R179 can be captured successfully, it means the Category element is not null, then MS-ASCNTC_R507 can be captured directly.
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R507");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R507
                    Site.CaptureRequirement(
                        507,
                        @"[In Category] A command response has a minimum of one Category element per Categories element.");

                    // If the validation is successful, then MS-ASCNTC_R181 can be captured. 
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R181");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R181
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        181,
                        @"[In Category] The value of this element[Category] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                    this.VerifyString();
                }
            }
        }

        /// <summary>
        /// This method is used to verify the Child element.
        /// </summary>
        /// <param name="childList">The array of Child element.</param>
        private void VerifyChild(string[] childList)
        {
            if (childList != null)
            {
                foreach (string child in childList)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R509");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R509
                    Site.CaptureRequirementIfIsNotNull(
                        child,
                        509,
                        @"[In Child] A command response has zero or more Child elements per Children element.");

                    // If the validation is successful, then MS-ASCNTC_R192 can be captured. 
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCNTC_R192");

                    // Verify MS-ASCNTC requirement: MS-ASCNTC_R192
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        192,
                        @"[In Child] The value of this element[Child] is a string data type, as specified in [MS-ASDTYPE] section 2.6.");

                    this.VerifyString();
                }
            }
        }
        #endregion

        #region Verify MS-ASDTYPE requirements
        /// <summary>
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainer()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // If the validation is successful, then MS-ASDTYPE_R9 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the string data type related requirements.
        /// </summary> 
        private void VerifyString()
        {
            // If the validation is successful, then MS-ASDTYPE_R88 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // If the validation is successful, then MS-ASDTYPE_R90 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R90
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                90,
                @"[In string Data Type] An element of this [string] type is declared as an element with a type attribute of ""string"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R91");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R91
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                91,
                @"[In string Data Type] Elements with a string data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // If the validation is successful, then MS-ASDTYPE_R94 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R94");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R94
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");
        }

        /// <summary>
        /// This method is used to verify the datetime data type related requirements.
        /// </summary>
        /// <param name="dateTime">The value of a datetime data type element.</param>
        private void VerifyDateTime(DateTime? dateTime)
        {
            // If the validation is successful, then MS-ASDTYPE_R12 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R12
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12,
                @"[In dateTime Data Type] It [dateTime]is declared as an element whose type attribute is set to ""dateTime"".");

            // If the value is not null, it means the date returned from server can be successfully converted to a DateTime type value, and the date should follow the following format, then requirement MS-ASDTYPE_R15 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R15
            Site.CaptureRequirementIfIsNotNull(
                dateTime,
                "MS-ASDTYPE",
                15,
                @"[In dateTime Data Type] All dates are given in Coordinated Universal Time (UTC) and are represented as a string in the following format.
YYYY-MM-DDTHH:MM:SS.MSSZ where
YYYY = Year (Gregorian calendar year)
MM = Month (01 - 12)
DD = Day (01 - 31)
HH = Number of complete hours since midnight (00 - 24)
MM = Number of complete minutes since start of hour (00 - 59)
SS = Number of seconds since start of minute (00 - 59)
MSS = Number of milliseconds");

            // If MS-ASDTYPE_R15 can be captured successfully, then MS-ASDTYPE_R16 can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R16");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R16
            Site.CaptureRequirementIfIsNotNull(
                dateTime,
                "MS-ASDTYPE",
                16,
                @"[In dateTime Data Type][in YYYY-MM-DDTHH:MM:SS.MSSZ ]The T serves as a separator, and the Z indicates that this time is in UTC.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R20");

            // ActiveSyncClient encoded dateTime data as inline strings, so if response is successfully returned this requirement can be verified.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R20
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                20,
                @"[In dateTime Data Type] Elements with a dateTime data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the integer data type related requirement.
        /// </summary>
        private void VerifyInteger()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R87");

            // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R87
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                87,
                @"[In integer Data Type] Elements with an integer data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the E-mail address related requirements.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        private void VerifyEmailAddress(string emailAddress)
        {
            // If the validation is successful, then MS-ASDTYPE_R99 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R99");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R99
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                99,
                @"[In E-Mail Address] An e-mail address is an unconstrained value of an element of the string type (section 2.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R100");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R100
            Site.CaptureRequirementIfIsTrue(
                RFC822AddressParser.IsValidAddress(emailAddress),
                "MS-ASDTYPE",
                100,
                @"[In E-Mail Address] However, a valid individual e-mail address MUST have the following format: ""local-part@domain"".");
        }
        #endregion

        #region Verify MS-ASWBXML requirements
        /// <summary>
        /// This method is used to verify MS-ASWBXML related requirements.
        /// </summary>
        private void VerifyWBXMLCapture()
        {
            // Get decoded data and capture requirement for decode processing
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (decodedData != null)
            {
                // Check out all tag-token
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    int codepage = decodeDataItem.Value;
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    bool isValidCodePage = codepage >= 0 && codepage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codepage);

                    // Capture the requirements in Contacts namespace
                    if (1 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R11");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R11
                        Site.CaptureRequirementIfAreEqual<string>(
                            "contacts",
                            codePageName.ToLower(CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            11,
                            @"[In Code Pages] [This algorithm supports] [Code page] 1 [that indicates] [XML namespace] Contacts.");

                        this.VerifyRequirementsRelateToCodePage1(codepage, tagName, token);
                    }

                    // Capture the requirements in Contacts2 namespace
                    if (12 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R22");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R22
                        Site.CaptureRequirementIfAreEqual<string>(
                            "contacts2",
                            codePageName.ToLower(CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            22,
                            @"[In Code Pages] [This algorithm supports] [Code page] 12 [that indicates] [XML namespace] Contacts2");

                        this.VerifyRequirementsRelateToCodePage12(codepage, tagName, token);
                    }
                }
            }
        }

        /// <summary>
        /// This method is used to verify the tags and tokens in WBXML code page 1.
        /// </summary>
        /// <param name="codePageNumber">The code page number.</param>
        /// <param name="tagName">The tag name that needs to be verified.</param>
        /// <param name="token">The token that needs to be verified.</param>
        private void VerifyRequirementsRelateToCodePage1(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Anniversary":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R73");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R73
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            73,
                            @"[In Code Page 1: Contacts] [Tag name] Anniversary [Token] 0x05");

                        break;
                    }

                case "AssistantName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R74");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R74
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            74,
                            @"[In Code Page 1: Contacts] [Tag name] AssistantName [Token] Anniversary 0x06");

                        break;
                    }

                case "AssistantPhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R75");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R75
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            75,
                            @"[In Code Page 1: Contacts] [Tag name] AssistantPhoneNumber [Token] 0x07");

                        break;
                    }

                case "Birthday":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R76");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R76
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            76,
                            @"[In Code Page 1: Contacts] [Tag name] Birthday [Token] 0x08");

                        break;
                    }

                case "Business2PhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R77");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R77
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            77,
                            @"[In Code Page 1: Contacts] [Tag name] Business2PhoneNumber [Token] 0x0C");

                        break;
                    }

                case "BusinessAddressCity":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R78");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R78
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            78,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessAddressCity [Token] 0x0D");

                        break;
                    }

                case "BusinessAddressCountry":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R79");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R79
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            79,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessAddressCountry [Token] 0x0E");

                        break;
                    }

                case "BusinessAddressPostalCode":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R80");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R80
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            80,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessAddressPostalCode [Token] 0x0F");

                        break;
                    }

                case "BusinessAddressState":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R81");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R81
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            81,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessAddressState [Token] 0x10");

                        break;
                    }

                case "BusinessAddressStreet":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R82");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R82
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            82,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessAddressStreet [Token] 0x11");

                        break;
                    }

                case "BusinessFaxNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R83");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R83
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            83,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessFaxNumber [Token] 0x12");

                        break;
                    }

                case "BusinessPhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R84");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R84
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            84,
                            @"[In Code Page 1: Contacts] [Tag name] BusinessPhoneNumber [Token] 0x13");

                        break;
                    }

                case "CarPhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R85");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R85
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            85,
                            @"[In Code Page 1: Contacts] [Tag name] CarPhoneNumber [Token] 0x14");

                        break;
                    }

                case "Categories":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R86");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R86
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x15,
                            token,
                            "MS-ASWBXML",
                            86,
                            @"[In Code Page 1: Contacts] [Tag name] Categories [Token] 0x15");

                        break;
                    }

                case "Category":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R87");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R87
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            87,
                            @"[In Code Page 1: Contacts] [Tag name] Category [Token] 0x16");

                        break;
                    }

                case "Children":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R88");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R88
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x17,
                            token,
                            "MS-ASWBXML",
                            88,
                            @"[In Code Page 1: Contacts] [Tag name] Children [Token] 0x17");

                        break;
                    }

                case "Child":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R89");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R89
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x18,
                            token,
                            "MS-ASWBXML",
                            89,
                            @"[In Code Page 1: Contacts] [Tag name] Child [Token] 0x18");

                        break;
                    }

                case "CompanyName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R90");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R90
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x19,
                            token,
                            "MS-ASWBXML",
                            90,
                            @"[In Code Page 1: Contacts] [Tag name] CompanyName [Token] 0x19");

                        break;
                    }

                case "Department":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R91");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R91
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1A,
                            token,
                            "MS-ASWBXML",
                            91,
                            @"[In Code Page 1: Contacts] [Tag name] Department [Token] 0x1A");

                        break;
                    }

                case "Email1Address":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R92");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R92
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1B,
                            token,
                            "MS-ASWBXML",
                            92,
                            @"[In Code Page 1: Contacts] [Tag name] Email1Address [Token] 0x1B");

                        break;
                    }

                case "Email2Address":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R93");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R93
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1C,
                            token,
                            "MS-ASWBXML",
                            93,
                            @"[In Code Page 1: Contacts] [Tag name] Email2Address [Token] 0x1C");

                        break;
                    }

                case "Email3Address":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R94");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R94
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1D,
                            token,
                            "MS-ASWBXML",
                            94,
                            @"[In Code Page 1: Contacts] [Tag name] Email3Address [Token] 0x1D");

                        break;
                    }

                case "FileAs":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R95");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R95
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1E,
                            token,
                            "MS-ASWBXML",
                            95,
                            @"[In Code Page 1: Contacts] [Tag name] FileAs [Token] 0x1E");

                        break;
                    }

                case "FirstName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R96");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R96
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1F,
                            token,
                            "MS-ASWBXML",
                            96,
                            @"[In Code Page 1: Contacts] [Tag name] FirstName [Token] 0x1F");

                        break;
                    }

                case "Home2PhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R97");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R97
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x20,
                            token,
                            "MS-ASWBXML",
                            97,
                            @"[In Code Page 1: Contacts] [Tag name] Home2PhoneNumber [Token] 0x20");

                        break;
                    }

                case "HomeAddressCity":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R98");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R98
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x21,
                            token,
                            "MS-ASWBXML",
                            98,
                            @"[In Code Page 1: Contacts] [Tag name] HomeAddressCity [Token] 0x21");

                        break;
                    }

                case "HomeAddressCountry":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R99");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R99
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x22,
                            token,
                            "MS-ASWBXML",
                            99,
                            @"[In Code Page 1: Contacts] [Tag name] HomeAddressCountry [Token] 0x22");

                        break;
                    }

                case "HomeAddressPostalCode":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R100");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R100
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x23,
                            token,
                            "MS-ASWBXML",
                            100,
                            @"[In Code Page 1: Contacts] [Tag name] HomeAddressPostalCode [Token] 0x23");

                        break;
                    }

                case "HomeAddressState":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R101");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R101
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x24,
                            token,
                            "MS-ASWBXML",
                            101,
                            @"[In Code Page 1: Contacts] [Tag name] HomeAddressState [Token] 0x24");

                        break;
                    }

                case "HomeAddressStreet":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R102");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R102
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x25,
                            token,
                            "MS-ASWBXML",
                            102,
                            @"[In Code Page 1: Contacts] [Tag name] HomeAddressStreet [Token] 0x25");

                        break;
                    }

                case "HomeFaxNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R103");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R103
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x26,
                            token,
                            "MS-ASWBXML",
                            103,
                            @"[In Code Page 1: Contacts] [Tag name] HomeFaxNumber [Token] 0x26");

                        break;
                    }

                case "HomePhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R104");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R104
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x27,
                            token,
                            "MS-ASWBXML",
                            104,
                            @"[In Code Page 1: Contacts] [Tag name] HomePhoneNumber [Token] 0x27");

                        break;
                    }

                case "JobTitle":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R105");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R105
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x28,
                            token,
                            "MS-ASWBXML",
                            105,
                            @"[In Code Page 1: Contacts] [Tag name] JobTitle [Token] 0x28");

                        break;
                    }

                case "LastName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R106");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R106
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x29,
                            token,
                            "MS-ASWBXML",
                            106,
                            @"[In Code Page 1: Contacts] [Tag name] LastName [Token] 0x29");

                        break;
                    }

                case "MiddleName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R107");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R107
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2A,
                            token,
                            "MS-ASWBXML",
                            107,
                            @"[In Code Page 1: Contacts] [Tag name] MiddleName [Token] 0x2A");

                        break;
                    }

                case "MobilePhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R108");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R108
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2B,
                            token,
                            "MS-ASWBXML",
                            108,
                            @"[In Code Page 1: Contacts] [Tag name] MobilePhoneNumber [Token] 0x2B");

                        break;
                    }

                case "OfficeLocation":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R109");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R109
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2C,
                            token,
                            "MS-ASWBXML",
                            109,
                            @"[In Code Page 1: Contacts] [Tag name] OfficeLocation [Token] 0x2C");

                        break;
                    }

                case "OtherAddressCity":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R110");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R110
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2D,
                            token,
                            "MS-ASWBXML",
                            110,
                            @"[In Code Page 1: Contacts] [Tag name] OtherAddressCity [Token] 0x2D");

                        break;
                    }

                case "OtherAddressCountry":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R111");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R111
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2E,
                            token,
                            "MS-ASWBXML",
                            111,
                            @"[In Code Page 1: Contacts] [Tag name] OtherAddressCountry [Token] 0x2E");

                        break;
                    }

                case "OtherAddressPostalCode":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R112");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R112
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x2F,
                            token,
                            "MS-ASWBXML",
                            112,
                            @"[In Code Page 1: Contacts] [Tag name] OtherAddressPostalCode [Token] 0x2F");

                        break;
                    }

                case "OtherAddressState":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R113");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R113
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x30,
                            token,
                            "MS-ASWBXML",
                            113,
                            @"[In Code Page 1: Contacts] [Tag name] OtherAddressState [Token] 0x30");

                        break;
                    }

                case "OtherAddressStreet":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R114");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R114
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x31,
                            token,
                            "MS-ASWBXML",
                            114,
                            @"[In Code Page 1: Contacts] [Tag name] OtherAddressStreet [Token] 0x31");

                        break;
                    }

                case "PagerNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R115");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R115
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x32,
                            token,
                            "MS-ASWBXML",
                            115,
                            @"[In Code Page 1: Contacts] [Tag name] PagerNumber [Token] 0x32");

                        break;
                    }

                case "RadioPhoneNumber":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R116");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R116
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x33,
                            token,
                            "MS-ASWBXML",
                            116,
                            @"[In Code Page 1: Contacts] [Tag name] RadioPhoneNumber [Token] 0x33");

                        break;
                    }

                case "Spouse":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R117");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R117
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x34,
                            token,
                            "MS-ASWBXML",
                            117,
                            @"[In Code Page 1: Contacts] [Tag name] Spouse [Token] 0x34");

                        break;
                    }

                case "Suffix":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R118");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R118
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x35,
                            token,
                            "MS-ASWBXML",
                            118,
                            @"[In Code Page 1: Contacts] [Tag name] Suffix [Token] 0x35");

                        break;
                    }

                case "Title":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R119");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R119
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x36,
                            token,
                            "MS-ASWBXML",
                            119,
                            @"[In Code Page 1: Contacts] [Tag name] Title [Token] 0x36");

                        break;
                    }

                case "WebPage":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R120");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R120
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x37,
                            token,
                            "MS-ASWBXML",
                            120,
                            @"[In Code Page 1: Contacts] [Tag name] WebPage [Token] 0x37");

                        break;
                    }

                case "YomiCompanyName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R121");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R121
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x38,
                            token,
                            "MS-ASWBXML",
                            121,
                            @"[In Code Page 1: Contacts] [Tag name] YomiCompanyName [Token] 0x38");

                        break;
                    }

                case "YomiFirstName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R122");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R122
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x39,
                            token,
                            "MS-ASWBXML",
                            122,
                            @"[In Code Page 1: Contacts] [Tag name] YomiFirstName [Token] 0x39");

                        break;
                    }

                case "YomiLastName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R123");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R123
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3A,
                            token,
                            "MS-ASWBXML",
                            123,
                            @"[In Code Page 1: Contacts] [Tag name] YomiLastName [Token] 0x3A");

                        break;
                    }

                case "Picture":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R124");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R124
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3C,
                            token,
                            "MS-ASWBXML",
                            124,
                            @"[In Code Page 1: Contacts] [Tag name] Picture [Token] 0x3C");

                        break;
                    }

                case "WeightedRank":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R127");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R127
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x3E,
                            token,
                            "MS-ASWBXML",
                            127,
                            @"[In Code Page 1: Contacts] [Tag name] WeightedRank<5> [Token] 0x3E");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There exists unexpected Tag in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 12.
        /// </summary>
        /// <param name="codePageNumber">The code page number.</param>
        /// <param name="tagName">The tag name that needs to be verified.</param>
        /// <param name="token">The token that needs to be verified.</param>
        private void VerifyRequirementsRelateToCodePage12(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "CustomerId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R531");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R531
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            531,
                            @"[In Code Page 12: Contacts2] [Tag name] CustomerId [Token] 0x05");

                        break;
                    }

                case "GovernmentId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R532");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R532
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            532,
                            @"[In Code Page 12: Contacts2] [Tag name] GovernmentId [Token] 0x06");

                        break;
                    }

                case "IMAddress":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R533");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R533
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            533,
                            @"[In Code Page 12: Contacts2] [Tag name] IMAddress[Token] 0x07");

                        break;
                    }

                case "IMAddress2":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R534");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R534
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            534,
                            @"[In Code Page 12: Contacts2] [Tag name] IMAddress2 [Token]0x08");

                        break;
                    }

                case "IMAddress3":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R535");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R535
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            535,
                            @"[In Code Page 12: Contacts2] [Tag name] IMAddress3 [Token] 0x09");

                        break;
                    }

                case "ManagerName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R536");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R536
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            536,
                            @"[In Code Page 12: Contacts2] [Tag name] ManagerName [Token] 0x0A");

                        break;
                    }

                case "CompanyMainPhone":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R537");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R537
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            537,
                            @"[In Code Page 12: Contacts2] [Tag name] CompanyMainPhone [Token] 0x0B");

                        break;
                    }

                case "AccountName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R538");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R538
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            538,
                            @"[In Code Page 12: Contacts2] [Tag name] AccountName [Token] 0x0C");

                        break;
                    }

                case "NickName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R539");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R539
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            539,
                            @"[In Code Page 12: Contacts2] [Tag name] NickName [Token] 0x0D");

                        break;
                    }

                case "MMS":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R540");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R540
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            540,
                            @"[In Code Page 12: Contacts2] [Tag name] MMS [Token] 0x0E");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There exists unexpected Tag in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }
        #endregion
    }
}