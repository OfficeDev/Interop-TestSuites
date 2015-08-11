namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using DataStructures = Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASAIRS.
    /// </summary>
    public partial class MS_ASAIRSAdapter : ManagedAdapterBase, IMS_ASAIRSAdapter
    {
        #region Verify Search command response
        /// <summary>
        /// This method is used to verify the Search command response relative requirements.
        /// </summary>
        /// <param name="response">Search command response.</param>
        /// <param name="store">A SearchStore object.</param>
        private void VerifySearchResponse(Microsoft.Protocols.TestSuites.Common.SearchResponse response, DataStructures.SearchStore store)
        {
            if (Common.IsRequirementEnabled(415, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R415");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R415
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    415,
                    @"[In Appendix B: Product Behavior] Implementation does check any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type, number of instances, order, and placement in the XML hierarchy, when receiving a Search command. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            this.VerifyCommonReqirements<Response.Search>(response);

            foreach (DataStructures.Search search in store.Results)
            {
                this.VerifyCommonElementsInResponse(search.Email);
            }
        }
        #endregion

        #region Verify Sync command response
        /// <summary>
        /// This method is used to verify the Sync command response relative requirements.
        /// </summary>
        /// <param name="response">Sync response.</param>
        /// <param name="syncStore">A SyncStore object.</param>
        private void VerifySyncResponse(SyncResponse response, DataStructures.SyncStore syncStore)
        {
            if (Common.IsRequirementEnabled(418, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R418");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R418
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    418,
                    @"[In Appendix B: Product Behavior] Implementation does check any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type, number of instances, order, and placement in the XML hierarchy, when receiving a Sync command. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R231");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R231
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                231,
                @"[In NativeBodyType] The NativeBodyType element is a required child element of the airsync:ApplicationData element ([MS-ASCMD]) in the Sync command that specifies the original format type of the item.");

            XmlDocument doc = new XmlDocument();
            if (!string.IsNullOrEmpty(response.ResponseDataXML))
            {
                doc.LoadXml(response.ResponseDataXML);
            }

            XmlNodeList applicationDataElements = doc.SelectNodes("//*[name()='ApplicationData']");
            foreach (XmlNode applicationDataElement in applicationDataElements)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R233. The count of NativeBodyType element per airsync:ApplicationData element is: {0}.", applicationDataElement.SelectNodes("*[name()='NativeBodyType']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R233
                Site.CaptureRequirementIfIsTrue(
                    applicationDataElement.SelectNodes("*[name()='NativeBodyType']").Count <= 1,
                    233,
                    @"[In NativeBodyType] A command response MUST have a maximum of one NativeBodyType element per airsync:ApplicationData element.");
            }

            if (syncStore != null)
            {
                if (syncStore.AddElements != null)
                {
                    foreach (DataStructures.Sync item in syncStore.AddElements)
                    {
                        if (item.Email.NativeBodyType != null)
                        {
                            string[] expecedValues = new string[] { "1", "2", "3" };
                            Common.VerifyActualValues("NativeBodyType", expecedValues, item.Email.NativeBodyType.ToString(), this.Site);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R236");

                            // Verify MS-ASAIRS requirement: MS-ASAIRS_R236
                            Site.CaptureRequirement(
                                236,
                                @"[In NativeBodyType] The following table defines the valid values [1, 2, 3] of the NativeBodyType enumeration.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R232");

                            // Verify MS-ASAIRS requirement: MS-ASAIRS_R232
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                232,
                                @"[In NativeBodyType] The value of this element [the NativeBodyType element] is an unsignedByte value ([MS-ASDTYPE] section 2.47).");

                            this.VerifyUnsignedByteDataType(item.Email.NativeBodyType);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R235");

                            // Verify MS-ASAIRS requirement: MS-ASAIRS_R235
                            Site.CaptureRequirementIfIsTrue(
                                this.activeSyncClient.ValidationResult,
                                235,
                                @"[In NativeBodyType] The NativeBodyType element MUST have no child elements.");
                        }

                        this.VerifyCommonElementsInResponse(item.Email);
                    }
                }
            }

            this.VerifyCommonReqirements<Response.Sync>(response);
        }
        #endregion

        #region Verify ItemOperations command response
        /// <summary>
        /// This method is used to verify the ItemOperations command response relative requirements.
        /// </summary>
        /// <param name="response">The ItemOperations command response.</param>
        /// <param name="itemOperationsStore">An ItemOperationsStore object.</param>
        private void VerifyItemOperationsResponse(ItemOperationsResponse response, DataStructures.ItemOperationsStore itemOperationsStore)
        {
            if (Common.IsRequirementEnabled(344, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R344");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R344
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    344,
                    @"[In Appendix B: Product Behavior] Implementation does check any of the XML elements specified in section 2.2.2 that are present in the command's XML body to ensure they comply with the requirements regarding data type, number of instances, order, and placement in the XML hierarchy, when receiving an ItemOperations command. (Exchange Server 2007 SP1 and above follow this behavior.)");
            }

            this.VerifyCommonReqirements<Response.ItemOperations>(response);

            foreach (DataStructures.ItemOperations itemOperation in itemOperationsStore.Items)
            {
                if (itemOperation.Email != null)
                {
                    this.VerifyCommonElementsInResponse(itemOperation.Email);
                }
            }
        }
        #endregion

        #region Verify common requirements in Search command, ItemOperations command and Sync command
        /// <summary>
        /// This method is used to verify the common requirements of Sync command, Search command and ItemOperations command response.
        /// </summary>
        /// <typeparam name="T">The type of response.</typeparam>
        /// <param name="response">The server response.</param>
        private void VerifyCommonReqirements<T>(ActiveSyncResponseBase<T> response)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R74");

            // If the schema validation is successful, then MS-ASAIRS_R74 can be captured.
            // Verify MS-ASAIRS requirement: MS-ASAIRS_R74
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                74,
                @"[In AllOrNone (BodyPreference)] The AllOrNone element MUST NOT be used in command responses.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R148");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R148
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                148,
                @"[In BodyPreference] A command response MUST NOT include the BodyPreference element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R297");

            // Verify MS-ASAIRS requirement: MS-ASAIRS_R297
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                297,
                @"[In TruncationSize (BodyPreference)] Command responses MUST NOT include the TruncationSize element.");

            if (Common.IsRequirementEnabled(412, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R412");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R412
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    412,
                    @"[In Appendix B: Product Behavior] Implementation does process commands[ItemOperations, Search, Sync] as specified in [MS-ASCMD]. (Exchange server 2007 SP1 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(413, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R413");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R413
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    413,
                    @"[In Appendix B: Product Behavior] Implementation does modify responses based on the elements specified in section 2.2.2 as specified for each element. (Exchange server 2007 SP1 and above follow this behavior.)");
            }

            XmlDocument doc = new XmlDocument();
            if (!string.IsNullOrEmpty(response.ResponseDataXML))
            {
                doc.LoadXml(response.ResponseDataXML);
            }

            XmlNodeList bodyElementNodes = doc.SelectNodes("//*[name()='Body']");
            foreach (XmlNode bodyElementNode in bodyElementNodes)
            {
                bool isDataValid = bodyElementNode.SelectNodes("*[name()='Data']").Count <= 1;

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R174. The count of Data element within each returned Body element is: {0}.", bodyElementNode.SelectNodes("*[name()='Data']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R174
                Site.CaptureRequirementIfIsTrue(
                    isDataValid,
                    174,
                    @"[In Data (Body)] A command response MUST have a maximum of one Data element within each returned Body element.");

                this.VerifyStringDataType();

                bool isEstimatedDataSizeValid = bodyElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R203. The count of EstimatedDataSize element per Body element, actually is: {0}.", bodyElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R203
                Site.CaptureRequirementIfIsTrue(
                    isEstimatedDataSizeValid,
                    203,
                    @"[In EstimatedDataSize (Body)] A command response MUST have a maximum of one EstimatedDataSize element per Body element.");

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    bool isPreviewElementValid = bodyElementNode.SelectNodes("*[name()='Preview']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R251. The count of Preview element per Body element is: {0}.", bodyElementNode.SelectNodes("*[name()='Preview']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R251
                    Site.CaptureRequirementIfIsTrue(
                        isPreviewElementValid,
                        251,
                        @"[In Preview (Body)] Command responses MUST have a maximum of one Preview element per Body element.");
                }

                bool isTruncatedElementValid = bodyElementNode.SelectNodes("*[name()='Truncated']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R280. The count of Truncated element per Body element is: {0}.", bodyElementNode.SelectNodes("*[name()='Truncated']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R280
                Site.CaptureRequirementIfIsTrue(
                    isTruncatedElementValid,
                    280,
                    @"[In Truncated (Body)] A command response MUST have a maximum of one Truncated element per Body element.");

                bool isTypeElementValid = bodyElementNode.SelectNodes("*[name()='Type']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R312. The count of Type element per Body element is: {0}.", bodyElementNode.SelectNodes("*[name()='Type']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R312
                Site.CaptureRequirementIfIsTrue(
                    isTypeElementValid,
                    312,
                    @"[In Type (Body)] A command response MUST have a maximum of one Type element per Body element.");
            }

            if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1")
                && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R59");

                // If the schema validation is successful, then MS-ASAIRS_R59 can be captured.
                // Verify MS-ASAIRS requirement: MS-ASAIRS_R59
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    59,
                    @"[In AllOrNone (BodyPartPreference)] The AllOrNone element MUST NOT be used in command responses.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R135");

                // If the schema validation is successful, then MS-ASAIRS_R135 can be captured.
                // Verify MS-ASAIRS requirement: MS-ASAIRS_R135
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    135,
                    @"[In BodyPartPreference] A command response MUST NOT include a BodyPartPreference element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R291");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R291
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    291,
                    @"[In TruncationSize (BodyPartPreference)] Command responses MUST NOT include the TruncationSize element.");

                XmlNodeList bodyPartElementNodes = doc.SelectNodes("//*[name()='BodyPart']");
                foreach (XmlNode bodyPartElementNode in bodyPartElementNodes)
                {
                    bool isDataValid = bodyPartElementNode.SelectNodes("*[name()='Data']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R183. The count of Data element within each returned BodyPart element is: {0}.", bodyPartElementNode.SelectNodes("*[name()='Data']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R183
                    Site.CaptureRequirementIfIsTrue(
                        isDataValid,
                        183,
                        @"[In Data (BodyPart)] A command response MUST have a maximum of one Data element within each returned BodyPart element.");

                    bool isEstimatedDataSizeValid = bodyPartElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R207. The count of EstimatedDataSize element per BodyPart element is: {0}.", bodyPartElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R207
                    Site.CaptureRequirementIfIsTrue(
                        isEstimatedDataSizeValid,
                        207,
                        @"[In EstimatedDataSize (BodyPart)] A command response MUST have a maximum of one EstimatedDataSize element per BodyPart element.");

                    bool isPreviewElementValid = bodyPartElementNode.SelectNodes("*[name()='Preview']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R257. The count of Preview element per BodyPart element is: {0}.", bodyPartElementNode.SelectNodes("*[name()='Preview']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R257
                    Site.CaptureRequirementIfIsTrue(
                        isPreviewElementValid,
                        257,
                        @"[In Preview (BodyPart)] Command responses MUST have a maximum of one Preview element per BodyPart element.");

                    bool isTruncatedElementValid = bodyPartElementNode.SelectNodes("*[name()='Truncated']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R286. The count of Truncated element per BodyPart element is: {0}.", bodyPartElementNode.SelectNodes("*[name()='Truncated']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R286
                    Site.CaptureRequirementIfIsTrue(
                        isTruncatedElementValid,
                        286,
                        @"[In Truncated (BodyPart)] A command response MUST have a maximum of one Truncated element per BodyPart element.");

                    bool isTypeElementValid = bodyPartElementNode.SelectNodes("*[name()='Type']").Count <= 1;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R317. The count of Type element per BodyPart element is: {0}.", bodyPartElementNode.SelectNodes("*[name()='Type']").Count);

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R317
                    Site.CaptureRequirementIfIsTrue(
                        isTypeElementValid,
                        317,
                        @"[In Type (BodyPart)] A command response MUST have a maximum of one Type element per BodyPart element.");
                }
            }

            XmlNodeList attachmentElementNodes = doc.SelectNodes("//*[name()='Attachment']");
            foreach (XmlNode attachmentElementNode in attachmentElementNodes)
            {
                bool isContentIdValid = attachmentElementNode.SelectNodes("*[name()='ContentId']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R164. The count of ContentId element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='ContentId']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R164
                Site.CaptureRequirementIfIsTrue(
                    isContentIdValid,
                    164,
                    @"[In ContentId] A command response MUST have a maximum of one ContentId element per Attachment element.");

                bool isContentLocationValid = attachmentElementNode.SelectNodes("*[name()='ContentLocation']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R169. The count of ContentLocation element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='ContentLocation']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R169
                Site.CaptureRequirementIfIsTrue(
                    isContentLocationValid,
                    169,
                    @"[In ContentLocation] A command response MUST have a maximum of one ContentLocation element per Attachment element.");

                bool isDisplayNamenValid = attachmentElementNode.SelectNodes("*[name()='DisplayName']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R192. The count of DisplayName element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='DisplayName']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R192
                Site.CaptureRequirementIfIsTrue(
                    isDisplayNamenValid,
                    192,
                    @"[In DisplayName] A command response MUST have a maximum of one DisplayName element per Attachment element.");

                bool isEstimatedDataSizeValid = attachmentElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R199. The count of EstimatedDataSize element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='EstimatedDataSize']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R199
                Site.CaptureRequirementIfIsTrue(
                    isEstimatedDataSizeValid,
                    199,
                    @"[In EstimatedDataSize (Attachment)] A command response MUST have a maximum of one EstimatedDataSize element per Attachment element.");

                bool isIsInLineValid = attachmentElementNode.SelectNodes("*[name()='IsInLine']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R218. The count of IsInline element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='IsInLine']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R218
                Site.CaptureRequirementIfIsTrue(
                    isIsInLineValid,
                    218,
                    @"[In IsInline] A command response MUST have a maximum of one IsInline element per Attachment element.");

                bool isMethodValid = attachmentElementNode.SelectNodes("*[name()='Method']").Count <= 1;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R222. The count of Method element per Attachment element is: {0}.", attachmentElementNode.SelectNodes("*[name()='Method']").Count);

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R222
                Site.CaptureRequirementIfIsTrue(
                    isMethodValid,
                    222,
                    @"[In Method] A command response MUST have a maximum of one Method element per Attachment element.");
            }
        }
        #endregion

        #region Verify common elements in Search command, ItemOperations command and Sync command
        /// <summary>
        /// This method is used to verify the common elements of Sync command, Search command and ItemOperations command response.
        /// </summary>
        /// <param name="email">An Email object.</param>
        private void VerifyCommonElementsInResponse(DataStructures.Email email)
        {
            if (email.Body != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R106");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R106
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    106,
                    @"[In Body] The Body element is a container ([MS-ASDTYPE] section 2.2).");

                this.VerifyContainerDataType();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R112");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R112
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    112,
                    @"[In Body] The Body element, if present, has the following required and optional child elements in this order [Type, EstimatedDataSize, Truncated, Data, Part, Preview].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R113");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R113
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    113,
                    @"[In Body] Type (section 2.2.2.22.1): This element is required.");

                if (email.Body.Data != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R177");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R177
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        177,
                        @"[In Data (Body)] In a response, the Data element MUST have no child elements.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R178");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R178
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        178,
                        @"[In Data (Body)] The content of the Data element is returned as a string in the format that is specified by the Type element (section 2.2.2.22.1).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R172");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R172
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        172,
                        @"[In Data] The value of this element [the Data element] is a string ([MS-ASDTYPE] section 2.6).");

                    this.VerifyStringDataType();
                }

                if (email.Body.Part != null)
                {
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        246,
                        @"[In Part] The value of this element [the Part element] is an integer ([MS-ASDTYPE] section 2.5).");

                    this.VerifyIntegerDataType();
                }

                if (email.Body.EstimatedDataSizeSpecified)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R204");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R204
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        204,
                        @"[In EstimatedDataSize (Body)] The EstimatedDataSize element MUST have no child elements.");
                }

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    if (email.Body.Preview != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R249");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R249
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            249,
                            @"[In Preview (Body)] The value of this element [the Preview element] is a string ([MS-ASDTYPE] section 2.6).");

                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R252");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R252
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            252,
                            @"[In Preview (Body)] The Preview element MUST have no child elements.");
                    }
                }

                if (email.Body.TruncatedSpecified)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R274");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R274
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        274,
                        @"[In Truncated] The value of this element [the Truncated element] is a boolean value ([MS-ASDTYPE] section 2.1) that specified whether the body or body part has been truncated.");

                    this.VerifyBooleanDataType(email.Body.Truncated);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R308");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R308
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    308,
                    @"[In Type (Body)] The Type element is a required element of the Body element (section 2.2.2.4).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R313");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R313
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    313,
                    @"[In Type (Body)] The Type element MUST have no child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R302");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R302
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    302,
                    @"[In Type] The value of this element[Type] is an unsignedByte value ([MS-ASDTYPE] section 2.7) that indicates the format type of the body content of the item.");

                this.VerifyUnsignedByteDataType(email.Body.Type);

                string[] expecedValues = new string[] { "1", "2", "3", "4" };
                Common.VerifyActualValues("Type", expecedValues, email.Body.Type.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R303");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R303
                Site.CaptureRequirement(
                    303,
                    @"[In Type] The following table defines the valid values [1, 2, 3, 4] of the Type element.");
            }

            if (email.BodyPart != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R121");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R121
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    121,
                    @"[In BodyPart] The BodyPart element is a container ([MS-ASDTYPE] section 2.2).");

                this.VerifyContainerDataType();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R124");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R124
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    124,
                    @"[In BodyPart] In a response, the airsync:ApplicationData element MUST be the parent element of the BodyPart element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R126");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R126
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    126,
                    @"[In BodyPart] The BodyPart element, if present, MUST have its required and optional child elements in the following order [Status, Type, EstimatedDataSize, Truncated, Data, Preview].");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R127");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R127
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    127,
                    @"[In BodyPart] Status (section 2.2.2.19). This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R269");

                // If MS-ASAIRS_R127 can be captured successfully, it means the Status element is a required element of the BodyPart element, then MS-ASAIRS_R269 can be captured directly.
                // Verify MS-ASAIRS requirement: MS-ASAIRS_R269
                Site.CaptureRequirement(
                    269,
                    @"[In Status] The Status element is a required child element of the BodyPart element (section 2.2.2.5) that indicates the success or failure of the response in returning Data element content (section 2.2.2.10.2) given the BodyPartPreference element settings (section 2.2.2.6) in the request.");

                string[] expecedValues = new string[] { "1", "176" };
                Common.VerifyActualValues("Status", expecedValues, email.BodyPart.Status.ToString(), this.Site);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R270");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R270
                Site.CaptureRequirement(
                    270,
                    @"[In Status] The following table lists valid values [1, 176] for the Status element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R128");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R128
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    128,
                    @"[In BodyPart] Type (section 2.2.2.22.2). This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R129");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R129
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    129,
                    @"[In BodyPart] EstimatedDataSize (section 2.2.2.12.3). This element is required.");

                if (email.BodyPart.Data != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R185");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R185
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        185,
                        @"[In Data (BodyPart)] In a response, the Data element MUST have no child elements.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R186");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R186
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        186,
                        @"[In Data (BodyPart)] The content of the Data element is returned as a string in the format that is specified by the Type element (section 2.2.2.22.2).");

                    this.VerifyStringDataType();
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R205");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R205
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                     205,
                     @"[In EstimatedDataSize (BodyPart)] The EstimatedDataSize element is a required child element of the BodyPart element (section 2.2.2.5).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R208");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R208
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    208,
                    @"[In EstimatedDataSize (BodyPart)] The EstimatedDataSize element MUST have no child elements.");

                if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                {
                    if (email.BodyPart.Preview != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R254");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R254
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            254,
                            @"[In Preview (BodyPart)] The value of this element [the Preview element] is a string ([MS-ASDTYPE] section 2.6).");

                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R258");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R258
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            258,
                            @"[In Preview (BodyPart)] The Preview element MUST have no child elements.");
                    }
                }

                if (email.BodyPart.TruncatedSpecified)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R274");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R274
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        274,
                        @"[In Truncated] The value of this element [the Truncated element] is a boolean value ([MS-ASDTYPE] section 2.1) that specified whether the body or body part has been truncated.");

                    this.VerifyBooleanDataType(email.BodyPart.Truncated);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R314");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R314
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    314,
                    @"[In Type (BodyPart)] The Type element is a required child element of the BodyPart element (section 2.2.2.5).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R318");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R318
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    318,
                    @"[In Type (BodyPart)] The Type element MUST have no child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R302");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R302
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    302,
                    @"[In Type] The value of this element[Type] is an unsignedByte value ([MS-ASDTYPE] section 2.7) that indicates the format type of the body content of the item.");

                this.VerifyUnsignedByteDataType(email.BodyPart.Type);
            }

            if (email.Attachments != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R102");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R102
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    102,
                  @"[In Attachments] The Attachments element has the following child elements: Attachment (section 2.2.2.2): At least one instance of this element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R87");

                // Verify MS-ASAIRS requirement: MS-ASAIRS_R87
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    87,
                    @"[In Attachment] The Attachment element is a required child element of the Attachments element (section 2.2.2.3).");

                foreach (Response.AttachmentsAttachment attachment in email.Attachments.Attachment)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R91");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R91
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        91,
                        @"[In Attachment] FileReference (section 2.2.2.13.1). This element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R211");

                    // If MS-ASAIRS_R91 can be captured, it means the FileReference element is a required child element of the Attachment element, so MS-ASAIRS_R211 can be captured directly after MS-ASAIRS_R91.
                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R211
                    Site.CaptureRequirement(
                        211,
                        @"[In FileReference (Attachment)] The FileReference element is a required child element of the Attachment element (section 2.2.2.2) that specifies the location of an item on the server to retrieve.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R210");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R210
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        210,
                        @"[In FileReference] The value of this element [the FileReference element] is a string value ([MS-ASDTYPE] section 2.6).");

                    this.VerifyStringDataType();

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R215");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R215
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        215,
                        @"[In FileReference (Fetch)] The FileReference element MUST have no child elements.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R92");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R92
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        92,
                        @"[In Attachment] Method (section 2.2.2.15). This element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R220");

                    // If MS-ASAIRS_R92 can be captured, it means the Method element is a required child element of the Attachment element, so MS-ASAIRS_R220 can be captured directly after MS-ASAIRS_R92.
                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R220
                    Site.CaptureRequirement(
                        220,
                        @"[In Method] The Method element is a required child element of the Attachment element (section 2.2.2.2) that identifies the method in which the attachment was attached.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R223");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R223
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        223,
                        @"[In Method] The Method element MUST have no child elements.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R221");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R221
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        221,
                        @"[In Method] The value of this element [the Method element] is an unsignedByte value ([MS-ASDTYPE] section 2.47).");

                    this.VerifyUnsignedByteDataType(attachment.Method);

                    string[] expecedValues = new string[] { "1", "2", "3", "4", "5", "6" };
                    Common.VerifyActualValues("Method", expecedValues, attachment.Method.ToString(), this.Site);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R224");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R224
                    Site.CaptureRequirement(
                        224,
                        @"[In Method] The following table defines the valid values [1, 2, 3, 4, 5, 6] of the Method element.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R93");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R93
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        93,
                        @"[In Attachment] EstimatedDataSize (section 2.2.2.12.1). This element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R198");

                    // If MS-ASAIRS_R93 can be captured, it means the EstimatedDataSize element is a required child element of the Attachment element, so MS-ASAIRS_R198 can be captured directly after MS-ASAIRS_R93.
                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R198
                    Site.CaptureRequirement(
                        198,
                        @"[In EstimatedDataSize (Attachment)] The EstimatedDataSize element is required child element of the Attachment element (section 2.2.2.2).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R195");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R195
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        195,
                        @"[In EstimatedDataSize] The value of this element [the EstimatedDataSize element] is an integer value ([MS-ASDTYPE] section 2.5).");

                    this.VerifyIntegerDataType();

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R200");

                    // Verify MS-ASAIRS requirement: MS-ASAIRS_R200
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        200,
                        @"[In EstimatedDataSize (Attachment)] The EstimatedDataSize element MUST have no child elements.");

                    if (attachment.ContentId != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R163");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R163
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            163,
                            @"[In ContentId] The value of this element [the ContentId element] is a string value ([MS-ASDTYPE] section 2.6).");

                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R165");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R165
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            165,
                            @"[In ContentId] The ContentId element MUST have no child elements.");
                    }

                    if (attachment.ContentLocation != null)
                    {
                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R168");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R168
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            168,
                            @"[In ContentLocation] The value of this element [ContentLocation] is a string ([MS-ASDTYPE] section 2.6) value.");

                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R170");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R170
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            170,
                            @"[In ContentLocation] The ContentLocation element MUST have no child elements.");
                    }

                    if (attachment.DisplayName != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R191");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R191
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            191,
                            @"[In DisplayName] The value of this element [the DisplayName element] is a string value ([MS-ASDTYPE] section 2.6).");

                        this.VerifyStringDataType();

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R193");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R193
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            193,
                            @"[In DisplayName] The DisplayName element MUST have no child elements.");
                    }

                    if (attachment.IsInlineSpecified)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R217");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R217
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            217,
                            @"[In IsInline] The value of this element [the IsInline element] is a boolean value ([MS-ASDTYPE] section 2.1).");
                        this.VerifyBooleanDataType(attachment.IsInline);

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASAIRS_R219");

                        // Verify MS-ASAIRS requirement: MS-ASAIRS_R219
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            219,
                            @"[In IsInline] The IsInline element MUST have no child elements.");
                    }
                }
            }
        }
        #endregion

        #region Verify requirements from MS-ASDTYPE
        /// <summary>
        /// This method is used to verify boolean data type related requirements.
        /// </summary>
        /// <param name="boolValue">The value of the Boolean element.</param>
        private void VerifyBooleanDataType(bool boolValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R4");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R4
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                4,
                @"[In boolean Data Type] It [a boolean] is declared as an element with a type attribute of ""boolean"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R5");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R5
            Site.CaptureRequirementIfIsTrue(
                Convert.ToInt32(boolValue).Equals(1) || Convert.ToInt32(boolValue).Equals(0),
                "MS-ASDTYPE",
                5,
                @"[In boolean Data Type] The value of a boolean element is an integer whose only valid values are 1 (TRUE) or 0 (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R7");

            // ActiveSyncClient encodes boolean data as inline strings, so if response is successfully returned this requirement can be verified.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R7
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                7,
                @"[In boolean Data Type] Elements with a boolean data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        /// <summary>
        /// This method is used to verify the integer data type related requirements.
        /// </summary>
        private void VerifyIntegerDataType()
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
        /// This method is used to verify the string data type related requirements.
        /// </summary>
        private void VerifyStringDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

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
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

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
        /// This method is used to verify the unsignedByte data type related requirements.
        /// </summary>
        /// <param name="byteValue">A byte value.</param>
        private void VerifyUnsignedByteDataType(byte? byteValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R123");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R123
            Site.CaptureRequirementIfIsTrue(
                (byteValue >= 0) && (byteValue <= 255),
                "MS-ASDTYPE",
                123,
                @"[In unsignedByte Data Type] The unsignedByte data type is an integer value between 0 and 255, inclusive.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R125");

            // If the schema validation is successful, then MS-ASDTYPE_R125 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R125
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                125,
                @"[In unsignedByte Data Type] Elements of this type [unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
        }
        #endregion

        #region Verify requirements from MS-ASWBXML
        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        private void VerifyWBXMLCapture()
        {
            // Get decode data and capture requirement for decode processing
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            if (decodedData != null)
            {
                // check out all tag-token
                foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    int codepage = decodeDataItem.Value;
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    bool isValidCodePage = codepage >= 0 && codepage <= 24;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codepage);

                    // Begin to capture requirement
                    if (17 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R27");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R27
                        Site.CaptureRequirementIfAreEqual<string>(
                            "airsyncbase",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            27,
                            @"[In Code Pages] [This algorithm supports] [Code page] 17 [that indicates] [XML namespace] AirSyncBase");

                        this.VerifyRequirementsRelateToCodePage17(codepage, tagName, token);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 17.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified. </param>
        private void VerifyRequirementsRelateToCodePage17(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Type":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R449");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R449
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            449,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Type [Token] 0x06");

                        break;
                    }

                case "Body":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R452");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R452
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            452,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Body [Token] 0x0A");

                        break;
                    }

                case "Data":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R453");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R453
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            453,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Data[Token]0x0B");

                        break;
                    }

                case "EstimatedDataSize":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R454");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R454
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            454,
                            @"[In Code Page 17: AirSyncBase] [Tag name] EstimatedDataSize [Token] 0x0C");

                        break;
                    }

                case "Truncated":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R455");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R455
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            455,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Truncated [Token] 0x0D");

                        break;
                    }

                case "Attachments":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R456");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R456
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            456,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Attachments [Token] 0x0E");

                        break;
                    }

                case "Attachment":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R457");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R457
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            457,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Attachment [Token] 0x0F");

                        break;
                    }

                case "DisplayName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R458");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R458
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            458,
                            @"[In Code Page 17: AirSyncBase] [Tag name] DisplayName [Token] 0x10");

                        break;
                    }

                case "FileReference":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R459");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R459
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            459,
                            @"[In Code Page 17: AirSyncBase] [Tag name] FileReference [Token] 0x11");

                        break;
                    }

                case "Method":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R460");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R460
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            460,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Method [Token] 0x12");

                        break;
                    }

                case "ContentId":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R461");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R461
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            461,
                            @"[In Code Page 17: AirSyncBase] [Tag name] ContentId [Token] 0x13");

                        break;
                    }

                case "ContentLocation":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R462");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R462
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            462,
                            @"[In Code Page 17: AirSyncBase] [Tag name] ContentLocation [Token] 0x14 (not used)");

                        break;
                    }

                case "IsInline":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R463");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R463
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x15,
                            token,
                            "MS-ASWBXML",
                            463,
                            @"[In Code Page 17: AirSyncBase] [Tag name] IsInline [Token] 0x15");

                        break;
                    }

                case "NativeBodyType":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R464");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R464
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            464,
                            @"[In Code Page 17: AirSyncBase] [Tag name] NativeBodyType [Token] 0x16");

                        break;
                    }

                case "ContentType":
                    break;

                case "Preview":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R466");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R466
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x18,
                            token,
                            "MS-ASWBXML",
                            466,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Preview<39> [Token] 0x18");

                        break;
                    }

                case "BodyPart":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R468");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R468
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1A,
                            token,
                            "MS-ASWBXML",
                            468,
                            @"[In Code Page 17: AirSyncBase] [Tag name] BodyPart<41> [Token] 0x1A");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R469");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R469
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1B,
                            token,
                            "MS-ASWBXML",
                            469,
                            @"[In Code Page 17: AirSyncBase] [Tag name] Status<42> [Token] 0x1B");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }
        #endregion
    }
}