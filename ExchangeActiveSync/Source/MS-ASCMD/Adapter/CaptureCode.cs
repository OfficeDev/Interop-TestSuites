namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.Response;
    using Microsoft.Protocols.TestSuites.MS_ASWBXML;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Partial class here contains server role capture code.
    /// </summary>
    public partial class MS_ASCMDAdapter : ManagedAdapterBase, IMS_ASCMDAdapter
    {
        #region Variables
        /// <summary>
        /// Give the WbxmlTracer Instance for WBXML data.
        /// </summary>
        private MS_ASWBXML msaswbxmlImplementation;

        /// <summary>
        /// A boolean that indicates whether the Class tag in WBXML code page 0 is exist.
        /// </summary>
        private bool isClassTagInPage0Exist = false;

        /// <summary>
        /// A boolean that indicates whether the Class tag in WBXML code page 6 is exist.
        /// </summary>
        private bool isClassTagInPage6Exist = false;

        #endregion

        #region CaptureTransportRequirements

        /// <summary>
        /// Verify the requirements of the transport when the response is received successfully.
        /// </summary>
        private void VerifyTransportRequirements()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R7");

            // This requirement can be captured when the response received successfully.
            Site.CaptureRequirement(
                7,
                @"[In Transport] All command (except the Autodiscover command) messages are encoded as WBXML.");
        }
        #endregion

        #region ASCMD_Element

        /// <summary>
        ///  Verify Message element when the Autodiscover response is received successfully.
        /// </summary>
        /// <param name="message">A received string of Message element.</param>
        private void VerifyMessageElement(string message)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(message, "The Message element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1896");

            // If the schema validation result is true and Message element is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1896,
                @"[In Message] Element Message in Autodiscover command response (section 2.2.2.1),the parent element is Error (section 2.2.3.60).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1897");

            // If the schema validation result is true and Message element is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1897,
                @"[In Message] None [Element Message in Autodiscover command response (section 2.2.2.1)has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1898");

            // If the schema validation result is true and Message element is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1898,
                @"[In Message] Element Message in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

            this.VerifyStringDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1899");

            // If the schema validation result is true and Message element is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1899,
                @"[In Message] Element Message in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");
        }

        /// <summary>
        ///  Verify DebugData element when the Autodiscover response is received successfully.
        /// </summary>
        /// <param name="debugData">A received string of DebugData element.</param>
        private void VerifyDebugDataElement(string debugData)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(debugData, "The DebugData element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1474");

            // If the schema validation result is true and DebugData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1474,
                @"[In DebugData] Element DebugData in Autodiscover command response (section 2.2.2.1), the parent element is Error (section 2.2.3.60).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1475");

            // If the schema validation result is true and DebugData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1475,
                @"[In DebugData] None [Element DebugData in Autodiscover command response (section 2.2.2.1) has no child element .]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1476");

            // If the schema validation result is true and DebugData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1476,
                @"[In DebugData] Element DebugData in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

            this.VerifyStringDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1477");

            // If the schema validation result is true and DebugData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1477,
                @"[In DebugData] Element DebugData in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");
        }

        /// <summary>
        ///  Verify ErrorCode element in Error element.
        /// </summary>
        /// <param name="errorCode">A received string of ErrorCode element.</param>
        private void VerifyErrorCodeElement(string errorCode)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(errorCode, "The ErrorCode element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1639");

            // If the schema validation result is true and ErrorCode is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1639,
                @"[In ErrorCode] Element ErrorCode in Autodiscover command response (section 2.2.2.1), the parent element is Error (section 2.2.3.60).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1640");

            // If the schema validation result is true and ErrorCode is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1640,
                @"[In ErrorCode] None [Element ErrorCode in Autodiscover command response (section 2.2.2.1) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1641");

            // If the schema validation result is true and ErrorCode is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1641,
                @"[In ErrorCode] Element ErrorCode in Autodiscover command response (section 2.2.2.1), the data type is integer ([MS-ASDTYPE] section 2.5).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1642");

            // If the schema validation result is true and ErrorCode is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1642,
                @"[In ErrorCode] Element ErrorCode in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");
        }

        /// <summary>
        /// Verify Status element for ResolveRecipients.
        /// </summary>
        /// <param name="status">The Status in ResolveRecipients.</param>
        private void VerifyStatusElementForResolveRecipients(int status)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4264");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4264,
                @"[In Status(ResolveRecipients)] The Status element is a required child element of the ResolveRecipients element, the Response element, the Availability element, the Certificates element, and the Picture element in ResolveRecipients command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2743");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2743,
                @"[In Status(ResolveRecipients)] Element Status in ResolveRecipients command response (section 2.2.2.14), the parent elements are ResolveRecipients (section 2.2.3.139), Response (section 2.2.3.143.5), Availability (section 2.2.3.16), Certificates (section 2.2.3.23.1), Picture (section 2.2.3.129.1).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2744");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2744,
                @"[In Status(ResolveRecipients)] None [Element Status in ResolveRecipients command response (section 2.2.2.14) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2746");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2746,
                @"[In Status(ResolveRecipients)] Element Status in ResolveRecipients command response (section 2.2.2.14), the number allowed is 1…1 (required).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2745");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2745,
                @"[In Status(ResolveRecipients)] Element Status in ResolveRecipients command response (section 2.2.2.14), the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            this.VerifyIntegerDataType();
        }

        /// <summary>
        /// Verify RecipientCount element for ResolveRecipients.
        /// </summary>
        /// <param name="recipientCount">The RecipientCount in ResolveRecipients.</param>
        private void VerifyRecipientCountElement(int recipientCount)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3765");

            // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3765,
                @"[In RecipientCount] The RecipientCount element is a required child element of the Response element and the Certificates element in ResolveRecipients command responses (section 2.2.2.14).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2435");

            // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2435,
                @"[In RecipientCount] Element RecipientCount in ResolveRecipients command response, the parent elements are Response (section 2.2.3.144.5), Certificates (section 2.2.3.23.1).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2436");

            // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2436,
                @"[In RecipientCount] None [Element RecipientCount in ResolveRecipients command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2437");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2437
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(int),
                recipientCount.GetType(),
                2437,
                @"[In RecipientCount] Element RecipientCount in ResolveRecipients command response, the data type is integer ([MS-ASDTYPE] section 2.6).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2438");

            // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2438,
                @"[In RecipientCount] Element RecipientCount in ResolveRecipients command response, the number allowed is 0…1 (optional).");
        }

        /// <summary>
        /// Verify Status element for ValidateCert.
        /// </summary>
        private void VerifyStatusElementForValidateCert()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4469");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4469,
                @"[In Status(ValidateCert)] The Status element is a required child element of the ValidateCert element and the Certficate element in ValidateCert command responses that indicates whether one or more certificates were successfully validated.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2771");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2771,
                @"[In Status(ValidateCert)] Element Status in ValidateCert command response (section 2.2.2.21), the parent element is ValidateCert (section 2.2.3.185), Certificate (section 2.2.3.19).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2772");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2772,
                @"[In Status(ValidateCert)] None [Element Status in ValidateCert command response (section 2.2.2.21) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2773");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2773,
                @"[In Status(ValidateCert)] Element Status in ValidateCert command response (section 2.2.2.21), the data type is unginedBynte ([MS-ASDTYPE] section 2.8).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2774");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2774,
                @"[In Status(ValidateCert)] Element Status in ValidateCert command response (section 2.2.2.21), the number allowed is 1...N (required).");
        }

        /// <summary>
        /// Verify Status element for Search.
        /// </summary>
        private void VerifyStatusElementForSearch()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4315");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4315,
                @"[In Status(Search)] The Status element is a required child element of the Search element, the Store element, and the gal:Picture element in Search command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2747");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2747,
                @"[In Status(Search)] Element Status in Search command response (section 2.2.2.15), the parent elements are Search (section 2.2.3.150), Store (section 2.2.3.168.2), gal:Picture (section 2.2.3.129.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2748");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2748,
                @"[In Status(Search)] None [Element Status in Search command response (section 2.2.2.15) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2749");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2749,
                @"[In Status(Search)] Element Status in Search command response (section 2.2.2.15), the data type is integer ([MS-ASDTYPE] section 2.6).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2750");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2750,
                @"[In Status(Search)] Element Status in Search command response (section 2.2.2.15), the number allowed is 1…1 (required).");
        }

        /// <summary>
        /// Verify Status element for Find.
        /// </summary>
        private void VerifyStatusElementForFind()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72171802");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72171802,
                @"[In Status (Find)] The Status element is a required child element of the Find element and the Response element in Find command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72171804");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72171804,
                @"[In Status (Find)] Element Status in Find command response (section 2.2.1.2), the parent element are Find (section 2.2.3.69),  Response (section 2.2.3.153.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72171805");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72171805,
                @"[In Status(Search)] None [Element Status in Search command response (section 2.2.2.15) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72171806");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72171806,
                @"[In Status (Find)] Element Status in Find command response (section 2.2.1.2), the data type is integer ([MS-ASDTYPE] section 2.6).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72171807");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72171807,
                @"[In Status (Find)] Element Status in Find command response (section 2.2.1.2), the number allowed is 1…1 (required).");
        }

        /// <summary>
        /// Verify Status element for Settings.
        /// </summary>
        private void VerifyStatusElementForSettings()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4382");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4382,
                @"[In Status(Settings)] The Status element is a required child element of the Settings element, the RightsManagementInformation element, the Oof element, the DevicePassword element, the DeviceInformation element, and the UserInformation element in Settings command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2755");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2755,
                @"[In Status(Settings)] Element Status in Settings command response, the parent elements are Settings (section 2.2.3.158.2), RightsManagementInformation (section 2.2.3.147), Oof (section 2.2.3.162), DevicePassword (section 2.2.3.46), DeviceInformation (section 2.2.3.45), UserInformation (section 2.2.3.182).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2756");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2756,
                @"[In Status(Settings)] None [Element Status in Settings command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2757");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2757,
                @"[In Status(Settings)] Element Status in Settings command response, the data type is integer ([MS-ASDTYPE] section 2.6).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2758");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2758,
                @"[In Status(Settings)] Element Status in Settings command response, the number allowed is 1 (required).");
        }

        /// <summary>
        /// Verify Status element for Sync command.
        /// </summary>
        private void VerifyStatusElementForSync()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4411");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4411,
                @"[In Status(Sync)] The Status element is a required child element of the the Collection element, the Change element, the Add element, and the Fetch element in Sync command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2767");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2767,
                @"[In Status(Sync)] Element Status in Sync command response (section 2.2.2.20), the parent elements are Collection (section 2.2.3.29.2), Change (section 2.2.3.24), Add (section 2.2.3.7.2), Delete (section 2.2.3.42.2), Fetch (section 2.2.3.63.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2768");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2768,
                @"[In Status(Sync)] None [Element Status in Sync command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2769");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2769,
                @"[In Status(Sync)] Element Status in Sync command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            this.VerifyIntegerDataType();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2770");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2770,
                @"[In Status(Sync)] Element Status in Sync command response, the number allowed is 1…1 (required).");
        }

        /// <summary>
        /// Verify ServerId element for Sync.
        /// </summary>
        /// <param name="serverId">The ServerId in Commands</param>
        private void VerifyServerIdElementForSync(string serverId)
        {
            if (!string.IsNullOrEmpty(serverId))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5701");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5701,
                    @"[In ServerId(Sync)] It [ServerId element] is a required child element of the Add element, the Change element, the Delete element, the Fetch element, and the SoftDelete element under the Commands element (section 2.2.3.32) in Sync command responses.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2605");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2605,
                    @"[In ServerId(Sync)] Element ServerId in Sync command response,the parent elements are Add (as a child of Commands) (section 2.2.3.7), Change (as a child of Commands), Fetch (as a child of Commands), Delete (as a child of Commands), (SoftDelete (as a child of Commands) (section 2.2.3.162).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2606");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2606,
                    @"[In ServerId(Sync)] None [Element ServerId in Sync command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2607");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2607,
                    @"[In ServerId(Sync)] Element ServerId in Sync command response, the data type is string,");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2608");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2608,
                    @"[In ServerId(Sync)] Element ServerId in Sync command response, the number allowed is 1…1 (required).");

                this.VerifyStringDataType();
            }
        }

        /// <summary>
        /// Verify ServerId element for Sync responses.
        /// </summary>
        /// <param name="serverId">The ServerId in Sync responses.</param>
        private void VerifyServerIdElementForSyncResponses(string serverId)
        {
            if (!string.IsNullOrEmpty(serverId))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5701");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5701,
                    @"[In ServerId(Sync)] It [ServerId element] is a required child element of the Add element, the Change element, the Delete element, the Fetch element, and the SoftDelete element under the Commands element (section 2.2.3.32) in Sync command responses.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5572");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5572,
                    @"[In ServerId(Sync)] Element ServerId in Sync command response, the parent elements are Add (as a child of Responses), Change (as a child of Responses), Delete (as a child of Responses), Fetch (as a child of Responses).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5573");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5573,
                    @"[In ServerId(Sync)] None [Element ServerId in Sync command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5574");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5574,
                    @"[In ServerId(Sync)]Element ServerId in Sync command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5575");

                // If the schema validation result is true and ServerId(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5575,
                    @"[In ServerId(Sync)] Element ServerId in Sync command response, the number allowed is 0…1 (optional).");

                this.VerifyStringDataType();
            }
        }

        /// <summary>
        /// Verify ApplicationData element for Add and Change in Sync command.
        /// </summary>
        private void VerifyApplicationDataForSyncAddChange()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5848");

            // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                5848,
                @"[In ApplicationData] The ApplicationData element is a required child element of the Change element, the Add element, and the Fetch element in Sync command responses.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1059");

            // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1059,
                @"[In ApplicationData] Element ApplicationData in Sync command response, the parent elements are Change (section 2.2.3.24), Add (section 2.2.3.7.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5533");

            // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                5533,
                @"[In ApplicationData] Element ApplicationData in Sync command response, the child element contains airsyncbase:Body (MS-ASAIRS] section 2.2.2.9), airsyncbase:BodyPart ([MS-ASAIRS] section 2.2.2.10), airsyncbase:NativeBodyType ([MS-ASAIRS] section 2.2.2.32), rm:RightsManagementLicense ([MS-ASRM] section 2.2.2.14).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1061");

            // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1061,
                @"[In ApplicationData]  Element ApplicationData in Sync command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1062");

            // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1062,
                @"[In ApplicationData]  Element ApplicationData in Sync command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();
        }

        /// <summary>
        /// Verify Forwardees element for Sync responses.
        /// </summary>
        /// <param name="forwardees">The Forwardees element for Sync responses.</param>
        private void VerifyForwardeesElementForSyncResponses()
        {
            #region Capture code for Forwardees
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66550801");

                // If the schema validation result is true and Forwardees is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66550801,
                    @"[In Forwardees] Element Forwardees in Sync command response (section 2.2.1.21), the parent element is MeetingRequest ([MS-ASEMAIL] section 2.2.2.48).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66550802");

                // If the schema validation result is true and Forwardees is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66550802,
                    @"[In Forwardees] None [Element Forwardees in Sync command response (section 2.2.1.21) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66550803");

                // If the schema validation result is true and Forwardees is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66550803,
                    @"[In Forwardees] Element Forwardees in Sync command response (section 2.2.1.21), the data type is container ([MS-ASDTYPE] section 2.2)");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66550804");

                // If the schema validation result is true and Forwardees is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66550804,
                    @"[In Forwardees] Element Forwardees in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();
                #endregion   
        }

        /// <summary>
        /// Verify Forwardee element for Sync responses.
        /// </summary>
        private void VerifyForwardeeElementForSyncResponses()
        {
            #region Capture code for Forwardee
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66250801");

                // If the schema validation result is true and Forwardee is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66250801,
                    @"[In Forwardee] Element Forwardee in Sync command response (section 2.2.1.21),  the parent element is Forwardees (section 2.2.3.79).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66250802");

                // If the schema validation result is true and Forwardee is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66250802,
                    @"[In Forwardee] Element Forwardee in Sync command response (section 2.2.1.21), the child elements are ForwardeeEmail (section 2.2.3.53), ForwardeeName (section 2.2.3.120.3).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66250803");

                // If the schema validation result is true and Forwardee is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66250803,
                    @"[In Forwardee] Element Forwardee in Sync command response (section 2.2.1.21), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66250804");

                // If the schema validation result is true and Forwardee is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    66250804,
                    @"[In Forwardee] Element Forwardee in Sync command response (section 2.2.1.21), the number allowed is 1…N (required).");

                this.VerifyContainerDataType();
                #endregion
        }
        
        /// <summary>
        /// Verify Forwardees Email element for Sync responses.
        /// </summary>
        private void VerifyForwardeeEmailElementForSyncResponses()
        {
            #region Capture code for Email

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64850802");

                // If the schema validation result is true and Email is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    64850802,
                    @"[In Email] [The Email element is a required child element of] the Forwardee element in Sync command responses that specifies the email address of the forwardee.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64850808");

                // If the schema validation result is true and Email is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    64850808,
                    @"[In Email] Element Email in Sync command response (section 2.2.1.21), the parent element is Forwardee (section 2.2.3.78).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64850809");

                // If the schema validation result is true and Email is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    64850809,
                    @"[In Email] None [Element Email in Sync command response (section 2.2.1.21) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64850810");

                // If the schema validation result is true and Email is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    64850810,
                    @"[In Email] Element Email in Sync command response (section 2.2.1.21), the data type is string ([MS-ASDTYPE] section 2.7).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64850811");

                // If the schema validation result is true and Email is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    64850811,
                    @"[In Email] Element Email in Sync command response (section 2.2.1.21), the number allowed is 1...1 (required).");

                this.VerifyStringDataType(); 
            #endregion
        }

        /// <summary>
        /// Verify Forwardees Name element for Sync responses.
        /// </summary>
        private void VerifyForwardeeNameElementForSyncResponses()
        {
            #region Capture code for Name
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651708");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                68651708,
                @"[In Name (SmartForward and Sync)] Element Name in Sync command response (section 2.2.1.21), the parent element is Forwardee (section 2.2.3.78).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651709");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                68651709,
                @"[In Name (SmartForward and Sync)] None [Element Name in Sync command response (section 2.2.1.21) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651710");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                68651710,
                @"[In Name (SmartForward and Sync)] Element Name in Sync command response (section 2.2.1.21), the data type is string ([MS-ASDTYPE] section 2.7).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651711");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                68651711,
                @"[In Name (SmartForward and Sync)] Element Name in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional).");

            this.VerifyStringDataType();
            #endregion
        }

        /// <summary>
        /// Verify ProposedEndTime Element when parent element is MeetingRequest for SyncResponses.
        /// </summary>
        private void VerifyMeetingRequestProposedEndTimeElementForSyncResponses()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901710");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901710,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the parent element is MeetingRequest ([MS-ASEMAIL] section 2.2.2.48).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901711");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901711,
                @"[In ProposedEndTime] None [Element ProposedEndTime in Sync command response (section 2.2.1.21) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901712");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901712,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the data type is Compact DateTime ([MS-ASDTYPE] section 2.7.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901713");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901713,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional)");

            VerifyDateTimeStructure();
        }

        /// <summary>
        /// Verify ProposedEndTime Element when parent element is Attendee for SyncResponses.
        /// </summary>
        private void VerifyAttendeeProposedEndTimeElementForSyncResponses()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901714");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901714,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the parent element is Attendee ([MS-ASCAL] section 2.2.2.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901715");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901715,
                @"[In ProposedEndTime] None [Element ProposedEndTime in Sync command response (section 2.2.1.21) has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901716");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901716,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the data type is Compact DateTime ([MS-ASDTYPE] section 2.7.2) .");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901717");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901717,
                @"[In ProposedEndTime] Element ProposedEndTime in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional).");

            VerifyDateTimeStructure();
        }

        /// <summary>
        /// Verify ProposedStartTime Element when parent element is MeetingRequest for SyncResponses.
        /// </summary>
        private void VerifyMeetingRequestProposedStartTimeElementForSyncResponses()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901731");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901731,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the parent element is MeetingRequest ([MS-ASEMAIL] section.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901732");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901732,
                @"[In ProposedStartTime] None [Element ProposedStartTime in Sync command response (section 2.2.1.21), has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901733");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901733,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the data type is Compact DateTime ([MS-ASDTYPE] .");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901734");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901734,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional).");

            VerifyDateTimeStructure();
        }

        /// <summary>
        /// Verify ProposedStartTime Element when parent element is Attendee for SyncResponses.
        /// </summary>
        private void VerifyAttendeeProposedStartTimeElementForSyncResponses()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901735");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901735,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the parent element is Attendee ([MS-ASCAL] section 2.2.2.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901736");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901736,
                @"[In ProposedStartTime] None [Element ProposedStartTime in Sync command response (section 2.2.1.21), has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901737");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901737,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the data type is Compact DateTime ([MS-ASDTYPE] section 2.7.2) .");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69901738");

            // If the schema validation result is true and Name is not null, this requirement can be verified.
            Site.CaptureRequirement(
                69901738,
                @"[In ProposedStartTime] Element ProposedStartTime in Sync command response (section 2.2.1.21), the number allowed is 0...1 (optional).");

            VerifyDateTimeStructure();
        }

        /// <summary>
        /// Verify child elements for Responses element in Sync command.
        /// </summary>
        /// <param name="element">The xml string of Response element.</param>
        private void VerifyElementsForResponses(XmlElement element)
        {
            if (element != null && element.HasChildNodes)
            {
                XmlNodeList addNodes = element.GetElementsByTagName("Add");
                XmlNodeList changeNodes = element.GetElementsByTagName("Change");
                XmlNodeList fetchNodes = element.GetElementsByTagName("Fetch");

                #region Capture code for Add
                if (addNodes.Count > 0)
                {
                    foreach (XmlNode addNode in addNodes)
                    {
                        if (addNode != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1043");

                            // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1043,
                                @"[In Add(Sync)] Element Add in Sync command response, the parent element is Responses (section 2.2.3.145).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1044");

                            // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1044,
                                @"[In Add(Sync)] Element Add in Sync command response, the child elements are ServerId, ClientId, Class, Status (section 2.2.3.167.16).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1045");

                            // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1045,
                                @"[In Add(Sync)] Element Add in Sync command response, the data type is container.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1046");

                            // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1046,
                                @"[In Add(Sync)] Element Add in Sync command response, the number allowed is 0...N (optional).");

                            this.VerifyContainerDataType();

                            bool hasStatus = false;
                            bool hasClientID = false;

                            foreach (XmlNode addChildNode in addNode.ChildNodes)
                            {
                                if (addChildNode.Name.Equals("Class", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1328");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R1328
                                    // If the schema validation result is true and Class is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1328,
                                        @"[In Class(Sync)] None [Element Class (Sync) in Sync command response has no child element.]");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1329");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R1329
                                    // If the schema validation result is true and Class is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1329,
                                        @"[In Class(Sync)] Element Class (Sync) in Sync command response, the data type is string.");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1330");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R1330
                                    // If the schema validation result is true and Class is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1330,
                                        @"[In Class(Sync)] Element Class (Sync) in Sync command response, the number allowed is 0...1 (optional).");

                                    // Add the debug information
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R13327");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R13327
                                    // If the schema validation result is true and Class is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        13327,
                                        @"[In Class(Sync)] Element Class (Sync) in Sync command response, the parent elements is Add (section 2.2.3.7.2).");
                                }

                                if (addChildNode.Name.Equals("Status", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    hasStatus = true;
                                    this.VerifyStatusElementForSync();
                                }

                                if (addChildNode.Name.Equals("ServerId", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    this.VerifyServerIdElementForSyncResponses(addChildNode.InnerXml);
                                }

                                if (addChildNode.Name.Equals("ClientId", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    hasClientID = true;

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R955");

                                    // If the schema validation result is true and ClientId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        955,
                                        @"[In ClientId (Sync)] The ClientId element is a required child element of the Add element in Sync command requests and responses that contains a unique identifier (typically an integer) that is generated by the client to temporarily identify a new object that is being created by using the Add element.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1347");

                                    // If the schema validation result is true and ClientId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1347,
                                        @"[In ClientId (Sync)] Element ClientId  in Sync command response, the parent element is Add (section 2.2.3.7.2).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1348");

                                    // If the schema validation result is true and ClientId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1348,
                                        @"[In ClientId (Sync)] None [Element ClientId  in Sync command response has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1349");

                                    // If the schema validation result is true and ClientId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1349,
                                        @"[In ClientId (Sync)] Element ClientId  in Sync command response, the data type is string.");

                                    this.VerifyStringDataType();

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1350");

                                    // If the schema validation result is true and ClientId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1350,
                                        @"[In ClientId (Sync)] Element ClientId  in Sync command response, the number allowed is 1…1 (required).");
                                }
                            }

                            Site.Assert.IsTrue(hasStatus, "The Status element in Add should not be null.");
                            Site.Assert.IsTrue(hasClientID, "The ClientId element in Add should not be null.");
                        }
                    }
                }
                #endregion

                #region Capture code for Change
                if (changeNodes.Count > 0)
                {
                    foreach (XmlNode changeNode in changeNodes)
                    {
                        if (changeNode != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1151");

                            // If the schema validation result is true and Change is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1151,
                                @"[In Change] Element Change in Sync command response, the parent element is Responses (section 2.2.3.145).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1152");

                            // If the schema validation result is true and Change is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1152,
                                @"[In Change] Element Change in Sync command response, the child elements are ServerId, Class, Status (section 2.2.3.167.16).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1153");

                            // If the schema validation result is true and Change is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1153,
                                @"[In Change] Element Change in Sync command response, the data type is container.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1154");

                            // If the schema validation result is true and Change is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1154,
                                @"[In Change] Element Change in Sync command response, the number allowed is 0...N (optional).");

                            this.VerifyContainerDataType();

                            bool hasStatus = false;

                            foreach (XmlNode changeChildNode in changeNode.ChildNodes)
                            {
                                if (changeChildNode.Name.Equals("Status", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    hasStatus = true;
                                    this.VerifyStatusElementForSync();
                                }

                                if (changeChildNode.Name.Equals("ServerId", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    this.VerifyServerIdElementForSyncResponses(changeChildNode.InnerXml);
                                }
                            }

                            Site.Assert.IsTrue(hasStatus, "The Status element in Change should not be null.");
                        }
                    }
                }
                #endregion

                #region Capture code for Fetch
                if (fetchNodes.Count > 0)
                {
                    foreach (XmlNode fetchNode in fetchNodes)
                    {
                        if (fetchNode != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1659");

                            // If the schema validation result is true and Fetch(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1659,
                                @"[In Fetch(Sync)] Element Fetch in Sync command response, the parent element is Responses (section 2.2.3.141).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1660");

                            // If the schema validation result is true and Fetch(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1660,
                                @"[In Fetch(Sync)] Element Fetch in Sync command response, the child elements are ServerId, Status (section 2.2.3.162.16), ApplicationData (section 2.2.3.11).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1661");

                            // If the schema validation result is true and Fetch(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1661,
                                @"[In Fetch(ItemOperations)] Element Fetch in Sync command response, the data type is container.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1662");

                            // If the schema validation result is true and Fetch(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1662,
                                @"[In Fetch(ItemOperations)] Element Fetch in Sync command response, the number allowed is  0...N (optional).");

                            this.VerifyContainerDataType();

                            bool hasStatus = false;

                            foreach (XmlNode fetchChildNode in fetchNode.ChildNodes)
                            {
                                if (fetchChildNode.Name.Equals("Status", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    hasStatus = true;
                                    this.VerifyStatusElementForSync();
                                }

                                if (fetchChildNode.Name.Equals("ServerId", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    this.VerifyServerIdElementForSyncResponses(fetchChildNode.InnerXml);
                                }

                                if (fetchChildNode.Name.Equals("ApplicationData", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    foreach (XmlNode applicationDataChildNode in fetchChildNode.ChildNodes)
                                    {
                                        if (applicationDataChildNode.Name.Equals("MeetingRequest", StringComparison.CurrentCultureIgnoreCase))
                                        {
                                            foreach (XmlNode meetingRequestChildNode in applicationDataChildNode.ChildNodes)
                                            {
                                                if (meetingRequestChildNode.Name.Equals("Forwardees", StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                    foreach (XmlNode forwardeesChildNode in meetingRequestChildNode.ChildNodes)
                                                    {
                                                        if (forwardeesChildNode.Name.Equals("Forwardee", StringComparison.CurrentCultureIgnoreCase))
                                                        {
                                                            if (forwardeesChildNode.InnerXml!=null)
                                                            {
                                                                VerifyForwardeeElementForSyncResponses();
                                                            }
                                                            foreach (XmlNode forwardeeChildNode in forwardeesChildNode.ChildNodes)
                                                            {
                                                                if (forwardeesChildNode.Name.Equals("ForwardeeEmail", StringComparison.CurrentCultureIgnoreCase))
                                                                {
                                                                    if (forwardeesChildNode!=null)
                                                                    {
                                                                        VerifyForwardeeEmailElementForSyncResponses();
                                                                    }
                                                                }

                                                                if (forwardeesChildNode.Name.Equals("ForwardeeName", StringComparison.CurrentCultureIgnoreCase))
                                                                {
                                                                    if (forwardeesChildNode != null)
                                                                    {
                                                                        VerifyForwardeeNameElementForSyncResponses();
                                                                    }
                                                                }
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }
                                    }

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5848");

                                    // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        5848,
                                        @"[In ApplicationData] The ApplicationData element is a required child element of the Change element, the Add element, and the Fetch element in Sync command responses.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1063");

                                    // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1063,
                                        @"[In ApplicationData] Element ApplicationData in Sync command response, the parent element is Fetch (section 2.2.3.63.2)");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1937");

                                    // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1937,
                                        @"[In ApplicationData] Element ApplicationData in Sync command response, the child elements are airsyncbase:Attachments ([MS-ASAIRS] section 2.2.28), airsyncbase:Body ([MS-ASAIRS] section 2.2.2.4), airsyncbase:NativeBodyType, rm:RightsManagementLicense.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1065");

                                    // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1065,
                                        @"[In ApplicationData] Element ApplicationData in Sync command response, the data type is container.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1066");

                                    // If the schema validation result is true and ApplicationData is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1066,
                                        @"[In ApplicationData] Element ApplicationData in Sync command response, the number allowed is  1…1 (required).");

                                    this.VerifyContainerDataType();
                                }
                            }

                            Site.Assert.IsTrue(hasStatus, "The Status element in Fetch should not be null.");
                        }
                    }
                }
                #endregion
            }
        }
        #endregion

        #region ASCMD_Command

        #region Capture code for Autodiscover command
        /// <summary>
        /// This method is used to verify the Auotdiscover response related requirements.
        /// </summary>
        /// <param name="autodiscoverResponse">Autodiscover command response.</param>
        private void VerifyAutodiscoverCommand(Microsoft.Protocols.TestSuites.Common.AutodiscoverResponse autodiscoverResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            if (autodiscoverResponse.ResponseData.Item is Response)
            {
                Response response = (Response)autodiscoverResponse.ResponseData.Item;
                #region Capture code for Autodiscover
                Site.Assert.IsNotNull(autodiscoverResponse.ResponseData, "Element Autodiscover in Autodiscover should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R8");

                // This requirement can be captured when the response received successfully.
                Site.CaptureRequirement(
                    8,
                    @"[In Transport] The Autodiscover command uses plain XML.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R41");

                // Verify the format of Autodiscover command of the request and responses messages, correct format is xml format.
                Site.CaptureRequirement(
                    41,
                    @"[In Autodiscover] The Autodiscover command request and response messages are sent in XML format, not WBXML format.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R806");

                // This requirement can be captured when the response received successfully, because ASClient use Http POST method to send out all MS-ASCMD commands in the implementation.
                Site.CaptureRequirement(
                    806,
                    @"[In Autodiscover] The Autodiscover element is a required element in Autodiscover command requests Responses that identifies the body of the HTTP POST as containing an Autodiscover command (section 2.2.2.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1083");

                // This requirement can be captured when the response received successfully.
                Site.CaptureRequirement(
                    1083,
                    @"[In Autodiscover] None [Element Autodiscover in Autodiscover command response has no parent element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1084");

                // This requirement can be captured when the response received successfully.
                Site.CaptureRequirement(
                    1084,
                    @"[In Autodiscover] Element Autodiscover in Autodiscover command response, the child element is Response (section 2.2.3.144.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1085");

                // This requirement can be captured when the response received successfully.
                Site.CaptureRequirement(
                    1085,
                    @"[In Autodiscover] Element Autodiscover in Autodiscover command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1086");

                // If the schema validation result is true and the Autodiscover element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1086,
                    @"[In Autodiscover] Element Autodiscover in Autodiscover command response, the number allowed is 1…1 (required).");

                this.VerifyContainerDataType();
                #endregion

                #region Capture code for Response
                Site.Assert.IsNotNull(response, "Element Response in Autodiscover command response should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3815");

                // If the schema validation result is true and the Response element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    3815,
                    @"[In Response(Autodiscover)] The Response element is a required child element of the Autodiscover element in Autodiscover command responses.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2479");

                // If the schema validation result is true and the Response element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2479,
                    @"[In Response(Autodiscover)] Element Response in Autodiscover command response (section 2.2.2.1), the parent element is Autodiscover (section 2.2.3.15).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2480");

                // If the schema validation result is true and the Response element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2480,
                    @"[In Response(Autodiscover)] Element Response in Autodiscover command response (section 2.2.2.1), the child elements are Culture (section 2.2.3.38), User (section 2.2.3.179), Action (section 2.2.3.6), Error (section 2.2.3.60).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2481");

                // If the schema validation result is true and the Response element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2481,
                    @"[In Response(Autodiscover)] Element Response in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2482");

                // If the schema validation result is true and the Response element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2482,
                    @"[In Response(Autodiscover)] Element Response in Autodiscover command response (section 2.2.2.1), the number allowed is 1...1 (required).");

                this.VerifyContainerDataType();
                #endregion

                #region Capture code for Culture
                if (response.Culture != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1458");

                    // If the schema validation result is true and the Culture element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1458,
                        @"[In Culture] Element Culture in Autodiscover command response (section 2.2.2.1), the parent element is Response (section 2.2.3.144.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1459");

                    // If the schema validation result is true and the Culture element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1459,
                        @"[In Culture] None [Element Culture in Autodiscover command response (section 2.2.2.1) has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1460");

                    // If the schema validation result is true and the Culture element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1460,
                        @"[In Culture] Element Culture in Autodiscover command response  (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1461");

                    // If the schema validation result is true and the Culture element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1461,
                        @"[In Culture] Element Culture in Autodiscover command response  (section 2.2.2.1), the number allowed is 0...1 (optional).");

                    this.VerifyStringDataType();
                }
                #endregion

                #region Capture code for User
                Site.Assert.IsNotNull(response.User, "Element User in Autodiscover command response should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4693");

                // If the schema validation result is true and the User element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    4693,
                    @"[In User] The User element is a required child element of the Response element in Autodiscover command responses that encapsulates information about the user to whom the Response element relates.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2887");

                // If the schema validation result is true and the User element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2887,
                    @"[In User] Element User in Autodiscover command response (section 2.2.2.1), the parent element is Response (section 2.2.3.144.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2888");

                // If the schema validation result is true and the User element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2888,
                    @"[In User] Element User in Autodiscover command response (section 2.2.2.1), the child elements are DisplayName (section 2.2.3.47.1), EMailAddress (section 2.2.3.52.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2889");

                // If the schema validation result is true and the User element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2889,
                    @"[In User] Element User in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2890");

                // If the schema validation result is true and the User element is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2890,
                    @"[In User] Element User in Autodiscover command response (section 2.2.2.1), the number allowed is 1...1 (required).");

                this.VerifyContainerDataType();
                #endregion

                #region Capture code for DisplayName(Autodiscover)
                if (response.User.DisplayName != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1518");

                    // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                    1518,
                    @"[In DisplayName(Autodiscover)] Element DisplayName in Autodiscover command response (section 2.2.2.1), the parent element is User (section 2.2.3.179).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1519");

                    // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                    1519,
                    @"[In DisplayName(Autodiscover)] None [Element DisplayName in Autodiscover command response (section 2.2.2.1)has no child element .]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1520");

                    // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                    1520,
                    @"[In DisplayName(Autodiscover)] Element DisplayName in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1521");

                    // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1521,
                        @"[In DisplayName(Autodiscover)] Element DisplayName in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                    this.VerifyStringDataType();
                }
                #endregion

                #region Capture code for EMailAddress
                Site.Assert.IsNotNull(response.User.EMailAddress, "The EMailAddress element should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2225");

                // If the schema validation result is true and EMailAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2225,
                    @"[In EMailAddress] The EMailAddress element is a required child element of the Request element in Autodiscover command requests and a required child element of the User element in Autodiscover command responses.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1583");

                // If the schema validation result is true and EMailAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1583,
                    @"[In EMailAddress] Element EMailAddress in Autodiscover command response, the parent element is User (section 2.2.3.173).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1584");

                // If the schema validation result is true and EMailAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1584,
                    @"[In EMailAddress] None [Element EMailAddress in Autodiscover command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1585");

                // If the schema validation result is true and EMailAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1585,
                    @"[In EMailAddress] Element EMailAddress in Autodiscover command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1586");

                // If the schema validation result is true and EMailAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1586,
                    @"[In EMailAddress] Element EMailAddress in Autodiscover command response, the number allowed is 1...1 (required).");
                #endregion

                #region Capture code for Action
                if (response.Action != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1027");

                    // If the schema validation result is true and Action is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1027,
                        @"[In Action] Element Action in Autodiscover command response (section 2.2.2.1), the parent element is Response (section 2.2.3.144.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1028");

                    // If the schema validation result is true and Action is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1028,
                        @"[In Action] Element Action in Autodiscover command response (section 2.2.2.1), the child elements are Redirect (section 2.2.3.138), Settings (section 2.2.3.158.1), Error (section 2.2.3.60).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1029");

                    // If the schema validation result is true and Action is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1029,
                        @"[In Action] Element Action in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1030");

                    // If the schema validation result is true and Action is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1030,
                        @"[In Action] Element Action in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                    this.VerifyContainerDataType();

                    #region Capture code for Settings(Autodiscover)
                    if (((Response)autodiscoverResponse.ResponseData.Item).Action.Settings != null && ((Response)autodiscoverResponse.ResponseData.Item).Action.Settings.Length > 0)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2623");

                        // If the schema validation result is true and Settings has any element, this requirement can be verified.
                        Site.CaptureRequirement(
                            2623,
                            @"[In Settings(Autodiscover)] Element Settings in Autodiscover command response (section 2.2.2.1), the parent element is Action (section 2.2.3.6).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2624");

                        // If the schema validation result is true and Settings has any element, this requirement can be verified.
                        Site.CaptureRequirement(
                            2624,
                            @"[In Settings(Autodiscover)] Element Settings in Autodiscover command response (section 2.2.2.1), the child element is Server (section 2.2.3.154).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2625");

                        // If the schema validation result is true and Settings has any element, this requirement can be verified.
                        Site.CaptureRequirement(
                            2625,
                            @"[In Settings(Autodiscover)] Element Settings in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2626");

                        // If the schema validation result is true and Settings has any element, this requirement can be verified.
                        Site.CaptureRequirement(
                            2626,
                            @"[In Settings(Autodiscover)] Element Settings in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                        this.VerifyContainerDataType();

                        #region Capture code for Server
                        foreach (ResponseActionServer autodiscoverResponseActionServer in ((Response)autodiscoverResponse.ResponseData.Item).Action.Settings)
                        {
                            Site.Assert.IsNotNull(autodiscoverResponseActionServer, "The Server element should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3890");

                            // If the schema validation result is true and Server element is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                3890,
                                @"[In Server] The Server element is a required child element of the Settings element in Autodiscover command responses that encapsulates settings that apply to a particular server in the Autodiscover command response.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2561");

                            // If the schema validation result is true and Server element is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2561,
                                @"[In Server] Element Server in Autodiscover command response (section 2.2.2.1), the parent element is Settings (section 2.2.3.158.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2562");

                            // If the schema validation result is true and Server element is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2562,
                                @"[In Server] Element Server in Autodiscover command response (section 2.2.2.1), the child elements are Type (section 2.2.3.176.1), Url (section 2.2.3.178), Name (section 2.2.3.114.1) ,ServerData (section 2.2.3.155).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2563");

                            // If the schema validation result is true and Server element is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2563,
                                @"[In Server] Element Server in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2564");

                            // If the schema validation result is true and Server element is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2564,
                                @"[In Server] Element Server in Autodiscover command response (section 2.2.2.1), the number allowed is 1...N (required).");

                            this.VerifyContainerDataType();

                            #region Capture code for Type(Autodiscover)
                            if (autodiscoverResponseActionServer.Type != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2859");

                                // If the schema validation result is true and Type element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2859,
                                    @"[In Type(Autodiscover)] Element Type in Autodiscover command response (section 2.2.2.1), the parent element is Server (section 2.2.3.154).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2860");

                                // If the schema validation result is true and Type element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2860,
                                    @"[In Type(Autodiscover)] None [Element Type in Autodiscover command response (section 2.2.2.1)has no child element .]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2861");

                                // If the schema validation result is true and Type element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2861,
                                    @"[In Type(Autodiscover)] Element Type in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2862");

                                // If the schema validation result is true and Type element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2862,
                                    @"[In Type(Autodiscover)] Element Type in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4648");

                                // Verify MS-ASCMD requirement: MS-ASCMD_R4648
                                Site.CaptureRequirementIfIsTrue(
                                    autodiscoverResponseActionServer.Type.ToString().Equals("MobileSync") || autodiscoverResponseActionServer.Type.ToString().Equals("CertEnroll"),
                                    4648,
                                    @"[In Type(Autodiscover)] The following are the valid values for the Type element[MobileSync, CertEnroll]:");

                                this.VerifyStringDataType();
                            }
                            #endregion

                            #region Capture code for Url
                            if (autodiscoverResponseActionServer.Url != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2883");

                                // If the schema validation result is true and Url element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2883,
                                    @"[In Url] Element Url in Autodiscover command response (section 2.2.2.1), the parent element is Server (section 2.2.3.154).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2884");

                                // If the schema validation result is true and Url element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2884,
                                    @"[In Url] None [Element Url in Autodiscover command response (section 2.2.2.1)has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2885");

                                // If the schema validation result is true and Url element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2885,
                                    @"[In Url] Element Url in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2886");

                                // If the schema validation result is true and Url element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2886,
                                    @"[In Url] Element Url in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                                this.VerifyStringDataType();
                            }
                            #endregion

                            #region Capture code for Name(Autodiscover)
                            if (autodiscoverResponseActionServer.Name != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1979");

                                // If the schema validation result is true and Name element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1979,
                                    @"[In Name(Autodiscover)] Element Name in Autodiscover command response (section 2.2.2.1), the parent element is Server (section 2.2.3.154).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1980");

                                // If the schema validation result is true and Name element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1980,
                                    @"[In Name(Autodiscover)] None [Element Name in Autodiscover command response (section 2.2.2.1)has  no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1981");

                                // If the schema validation result is true and Name element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1981,
                                    @"[In Name(Autodiscover)] Element Name in Autodiscover command response (section 2.2.2.1), the data type is string ([MS-ASDTYPE] section 2.7).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1982");

                                // If the schema validation result is true and Name element is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1982,
                                    @"[In Name(Autodiscover)] Element Name in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                                this.VerifyStringDataType();
                            }
                            #endregion
                        }
                        #endregion
                    }
                    #endregion
                    #endregion

                    #region Capture Code for Error in Action
                    if (((Response)autodiscoverResponse.ResponseData.Item).Action.Error != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1631");

                        // If the Error element is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1631,
                            @"[In Error] Element Error in Autodiscover command response (section 2.2.2.1), the parent element is Action (section 2.2.3.6).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1632");

                        // If the schema validation result is true and Error element is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1632,
                            @"[In Error] Element Error in Autodiscover command response (section 2.2.2.1), the child elements are Status (section 2.2.3.162.1), Message (section 2.2.3.98), DebugData (section 2.2.3.40), ErrorCode (section 2.2.3.61).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1633");

                        // If the schema validation result is true and Error element is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1633,
                            @"[In Error] Element Error in Autodiscover command response (section 2.2.2.1), the data type is container ([MS-ASDTYPE] section 2.2).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1634");

                        // If the schema validation result is true and Error element is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1634,
                            @"[In Error] Element Error in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                        #region Capture code for Status(Autodiscover)
                        if (((Response)autodiscoverResponse.ResponseData.Item).Action.Error.Status != null)
                        {
                            int status;

                            Site.Assert.IsTrue(int.TryParse(((Response)autodiscoverResponse.ResponseData.Item).Action.Error.Status, out status), "The Status element should be integer.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2693");

                            // If the schema validation result is true, Status is not null and is an integer, this requirement can be verified.
                            Site.CaptureRequirement(
                                2693,
                                @"[In Status(Autodiscover)] Element Status in Autodiscover command response (section 2.2.2.1), the data type is integer ([MS-ASDTYPE] section 2.6).");

                            this.VerifyIntegerDataType();

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2691");

                            // If the schema validation result is true and Status is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2691,
                                @"[In Status(Autodiscover)] Element Status in Autodiscover command response (section 2.2.2.1), the parent element is Error (section 2.2.3.60).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2692");

                            // If the schema validation result is true and Status is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2692,
                                @"[In Status(Autodiscover)] None [Element Status in Autodiscover command response (section 2.2.2.1) has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2694");

                            // If the schema validation result is true and Status is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2694,
                                @"[In Status(Autodiscover)] Element Status in Autodiscover command response (section 2.2.2.1), the number allowed is 0...1 (optional).");

                            Common.VerifyActualValues("Status(Autodiscover)", AdapterHelper.ValidStatus(new string[] { "1", "2" }), ((Response)autodiscoverResponse.ResponseData.Item).Action.Error.Status, this.Site);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3998");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R3998
                            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                            Site.CaptureRequirement(
                                3998,
                                @"[In Status(Autodiscover)] The following table lists the status codes [1, 2] for the Autodiscover command. For status values common to all ActiveSync commands, see section 2.2.4.");
                        }
                        #endregion

                        if (((Response)autodiscoverResponse.ResponseData.Item).Action.Error.Message != null)
                        {
                            this.VerifyMessageElement(((Response)autodiscoverResponse.ResponseData.Item).Action.Error.Message);
                        }

                        if (((Response)autodiscoverResponse.ResponseData.Item).Action.Error.DebugData != null)
                        {
                            this.VerifyDebugDataElement(((Response)autodiscoverResponse.ResponseData.Item).Action.Error.DebugData);
                        }

                        if (((Response)autodiscoverResponse.ResponseData.Item).Action.Error.ErrorCode != null)
                        {
                            this.VerifyErrorCodeElement(((Response)autodiscoverResponse.ResponseData.Item).Action.Error.ErrorCode);
                        }
                    }
                    #endregion
                }
                #endregion
            }
            else if (autodiscoverResponse.ResponseData.Item is Microsoft.Protocols.TestSuites.Common.Response.AutodiscoverResponse)
            {
                #region Capture code for Error
                Microsoft.Protocols.TestSuites.Common.Response.AutodiscoverResponse response = (Microsoft.Protocols.TestSuites.Common.Response.AutodiscoverResponse)autodiscoverResponse.ResponseData.Item;
                if (response.Error != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1635");

                    // If the schema validation result is true and Error element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1635,
                        @"[In Error] Element Error in Autodiscover command response, the parent element is Response (section 2.2.3.140.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1636");

                    // If the schema validation result is true and Error element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1636,
                        @"[In Error] Element Error in Autodiscover command response, the child elements are ErrorCode, Message, DebugData.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1637");

                    // If the schema validation result is true and Error element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1637,
                        @"[In Error] Element Error in Autodiscover command response, the data type is container.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1638");

                    // If the schema validation result is true and Error element is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1638,
                        @"[In Error] Element Error in Autodiscover command response, the number allowed is 0...1 (optional).");

                    if (response.Error.Time != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5698");

                        // If the schema validation result is true and Time attribute is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5698,
                            @"[In Error] The value of the attribute Time is a string that represents the time of day in that the request that generated the response was submitted.");

                        this.VerifyTimeAttribute(((Microsoft.Protocols.TestSuites.Common.Response.AutodiscoverResponse)autodiscoverResponse.ResponseData.Item).Error.Time);
                    }

                    if (response.Error.IdSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5545");

                        // If the schema validation result is true and Id attribute is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5545,
                            @"[In Error] The value of the attribute Id is an unsigned integer that uniquely identifies the server (within its domain) that generated the response.");
                    }

                    if (response.Error.ErrorCode != null)
                    {
                        this.VerifyErrorCodeElement(response.Error.ErrorCode);
                    }

                    if (response.Error.Message != null)
                    {
                        this.VerifyMessageElement(response.Error.Message);
                    }

                    if (((Microsoft.Protocols.TestSuites.Common.Response.AutodiscoverResponse)autodiscoverResponse.ResponseData.Item).Error.DebugData != null)
                    {
                        this.VerifyDebugDataElement(response.Error.DebugData);
                    }
                }
                #endregion
            }
        }

        /// <summary>
        /// This method is used to verify the ABNF syntax of Time attribute in Error Element 
        /// </summary>
        /// <param name="time">The time string.</param>
        private void VerifyTimeAttribute(string time)
        {
            string[] timeArr = time.Split('.');
            Site.Assert.AreEqual<int>(
                2,
                timeArr.Length,
                @"The time should be split by . as two parts.");
            string fractionalSeconds = timeArr[1];
            Site.Assert.AreEqual<int>(
                7,
                fractionalSeconds.Length,
                @"Fractional seconds should take seven decimal places");

            string[] subTimeArr = timeArr[0].Split(':');
            Site.Assert.AreEqual<int>(
                3,
                subTimeArr.Length,
                @"There should be three values for hours, minutes and seconds");

            string seconds = subTimeArr[2];

            Site.Assert.AreEqual<int>(
                2,
                seconds.Length,
                @"Seconds {0} should take two decimal places", seconds);

            int secondsValue = Convert.ToInt32(seconds);

            Site.Assert.IsTrue(
                secondsValue >= 0 && secondsValue <= 59,
                "The value for seconds should be valid.");

            string minutes = subTimeArr[1];

            Site.Assert.AreEqual<int>(
                2,
                minutes.Length,
                @"Minutes {0} should take two decimal places", minutes);

            int minutesValue = Convert.ToInt32(minutes);

            Site.Assert.IsTrue(
                minutesValue >= 0 && minutesValue <= 59,
                "The value for seconds should be valid.");

            string hours = subTimeArr[0];

            Site.Assert.AreEqual<int>(
                2,
                hours.Length,
                @"Hours {0} should take two decimal places", hours);

            int hoursValue = Convert.ToInt32(hours);

            Site.Assert.IsTrue(
                hoursValue >= 0 && hoursValue <= 23,
                "The value for seconds should be valid.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5547");

            // If the assertions above all pass, this requirement can be verified.
            Site.CaptureRequirement(
                5547,
                @"[In Error] [The ABNF syntax of the value of the Time attribute is] 
                time_val            = hours "":"" minutes "":"" seconds [""."" fractional_seconds]

                hours               = 2*DIGIT  ; 00 - 23, representing a 24-hour clock
                minutes             = 2*DIGIT  ; 00 - 59
                seconds             = 2*DIGIT  ; 00 - 59
                fractional_seconds  = 7*DIGIT  ; fractional seconds, always to 7 decimal places");
        }

        #endregion

        #region Capture code for FolderCreate command
        /// <summary>
        /// This method is used to verify the FolderCreate response related requirements.
        /// </summary>
        /// <param name="folderCreateResponse">FolderCreate command response.</param>
        private void VerifyFolderCreateCommand(FolderCreateResponse folderCreateResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(folderCreateResponse.ResponseData, "The FolderCreate element should not be null.");

            #region Capture code for FolderCreate
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3084");

            // If the schema validation result is true and FolderCreate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3084,
                @"[In FolderCreate] The FolderCreate element is a required element in FolderCreate command requests and FolderCreate command responses that identifies the body of the HTTP POST as containing a FolderCreate command (section 2.2.2.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1683");

            // If the schema validation result is true and FolderCreate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1683,
                @"[In FolderCreate] None [Element FolderCreate in FolderCreate command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1684");

            // If the schema validation result is true and FolderCreate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1684,
                @"[In FolderCreate] Element FolderCreate in FolderCreate command response, the child elements are SyncKey, ServerId (section 2.2.3.151.1), Status (section 2.2.3.162.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1685");

            // If the schema validation result is true and FolderCreate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1685,
                @"[In FolderCreate] Element FolderCreate in FolderCreate command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1686");

            // If the schema validation result is true and FolderCreate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1686,
                @"[In FolderCreate] Element FolderCreate in FolderCreate command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for SyncKey(FolderCreate)
            if (folderCreateResponse.ResponseData.SyncKey != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2803");

                // If the schema validation result is true and SyncKey(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2803,
                    @"[In SyncKey(FolderCreate)] Element SyncKey in FolderCreate command response, the parent element is FolderCreate.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2804");

                // If the schema validation result is true and SyncKey(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2804,
                    @"[In SyncKey(FolderCreate] None [Element SyncKey in FolderCreate command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2805");

                // If the schema validation result is true and SyncKey(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2805,
                    @"[In SyncKey(FolderCreate)] Element SyncKey in FolderCreate command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2806");

                // If the schema validation result is true and SyncKey(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2806,
                    @"[In SyncKey(FolderCreate)] Element SyncKey in FolderCreate command response, the number allowed is 0...1 (optional).");

                this.VerifyStringDataType();
            }
            #endregion

            #region Capture code for ServerId(FolderCreate)
            if (folderCreateResponse.ResponseData.ServerId != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2569");

                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2569,
                    @"[In ServerId(FolderCreate)] Element ServerId in FolderCreate command response,the parent element is FolderCreate (section 2.2.3.67).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2570");

                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2570,
                    @"[In ServerId(FolderCreate)] None [Element ServerId in FolderCreate command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2571");

                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2571,
                    @"[In ServerId(FolderCreate)] Element ServerId in FolderCreate command response, the data type is string ([MS-ASDTYPE] section 2.7).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2572");

                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2572,
                    @"[In ServerId(FolderCreate)] Element ServerId in FolderCreate command response, the number allowed is 0...1 (optional).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5872");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5872
                Site.CaptureRequirementIfIsTrue(
                    folderCreateResponse.ResponseData.ServerId.Length <= 64,
                    5872,
                    @"[In ServerId(FolderCreate)] The ServerId element value is not larger than 64 characters in length.");

                this.VerifyStringDataType();
            }
            #endregion

            #region Capture code for Status(FolderCreate)
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2695");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2695,
                @"[In Status(FolderCreate)] Element Status in FolderCreate command response, the parent element is FolderCreate (section 2.2.3.67).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2696");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2696,
                @"[In Status(FolderCreate)] None [Element Status in FolderCreate command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2697");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2697,
                @"[In Status(FolderCreate)] Element Status in FolderCreate command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2698");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2698,
                @"[In Status(FolderCreate)] Element Status in FolderCreate command response, the number allowed is 1…1 (required).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4005");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4005,
                @"[In Status(FolderCreate)] The Status element is a required child element of the FolderCreate element in FolderCreate command responses that indicates the success or failure of a FolderCreate command request (section 2.2.2.2).");

            Common.VerifyActualValues("Status(FolderCreate)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "5", "6", "9", "10", "11", "12" }), folderCreateResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4008");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4008
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4008,
                @"[In Status(FolderCreate)] The following table lists the status codes [1,2,3,5,6,9,10,11,12] for the FolderCreate command (section 2.2.2.2). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

            this.VerifyIntegerDataType();
            #endregion
        }
        #endregion

        #region Capture code for FolderDelete command
        /// <summary>
        /// This method is used to verify the FolderDelete response related requirements.
        /// </summary>
        /// <param name="folderDeleteResponse">FolderDelete command response.</param>
        private void VerifyFolderDeleteCommand(FolderDeleteResponse folderDeleteResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(folderDeleteResponse.ResponseData, "The FolderDelete element should not be null.");

            #region Capture code for FolderDelete
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3087");

            // If the schema validation result is true and FolderDelete is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3087,
                @"[In FolderDelete] The FolderDelete element is a required element in FolderDelete command requests and FolderDelete command responses that identifies the body of the HTTP POST as containing a FolderDelete command (section 2.2.2.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1691");

            // If the schema validation result is true and FolderDelete is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1691,
                @"[In FolderDelete] None [Element FolderDelete in FolderDelete command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1692");

            // If the schema validation result is true and FolderDelete is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1692,
                @"[In FolderDelete] Element FolderDelete in FolderDelete command response, the child elements are  SyncKey, Status (section 2.2.3.162.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1693");

            // If the schema validation result is true and FolderDelete is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1693,
                @"[In FolderDelete] Element FolderDelete in FolderDelete command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1694");

            // If the schema validation result is true and FolderDelete is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1694,
                @"[In FolderDelete] Element FolderDelete in FolderDelete command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for SyncKey(FolderDelete)
            if (folderDeleteResponse.ResponseData.SyncKey != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2811");

                // If the schema validation result is true and SyncKey(FolderDelete) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2811,
                    @"[In SyncKey(FolderDelete)] Element SyncKey in FolderDelete command response, the parent element is FolderDelete.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2812");

                // If the schema validation result is true and SyncKey(FolderDelete) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2812,
                    @"[In SyncKey(FolderDelete)] None [Element SyncKey in FolderDelete command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2813");

                // If the schema validation result is true and SyncKey(FolderDelete) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2813,
                    @"[In SyncKey(FolderDelete)] Element SyncKey in FolderDelete command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2814");

                // If the schema validation result is true and SyncKey(FolderDelete) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2814,
                    @"[In SyncKey(FolderDelete)] Element SyncKey in FolderDelete command response, the number allowed is 0...1 (optional).");

                this.VerifyStringDataType();
            }
            #endregion

            #region Capture code for Status(FolderDelete)
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4039");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4039,
                @"[In Status(FolderDelete)] The Status element is a required child element of the FolderDelete element in FolderDelete command responses that indicates the success or failure of the FolderDelete command request (section 2.2.2.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2699");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2699,
                @"[In Status(FolderDelete)] Element Status in FolderDelete command response, the parent element is FolderDelete (section 2.2.3.68).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2700");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2700,
                @"[In Status(FolderDelete)] None [Element Status in FolderDelete command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2701");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2701,
                @"[In Status(FolderDelete)] Element Status in FolderDelete command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2702");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2702,
                @"[In Status(FolderDelete)] Element Status in FolderDelete command response, the number allowed is 1…1 (required).");

            Common.VerifyActualValues("Status(FolderDelete)", AdapterHelper.ValidStatus(new string[] { "1", "3", "4", "6", "9", "10", "11" }), folderDeleteResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4042");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4042
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4042,
                @"[In Status(FolderDelete)] The following table lists the status codes [1,3,4,6,9,10,11] for the FolderDelete command (section 2.2.2.3). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

            this.VerifyIntegerDataType();
            #endregion
        }
        #endregion

        #region Capture code for FolderSync command
        /// <summary>
        /// This method is used to verify the FolderSync response related requirements.
        /// </summary>
        /// <param name="folderSyncResponse">FolderSync command response.</param>
        private void VerifyFolderSyncCommand(FolderSyncResponse folderSyncResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(folderSyncResponse.ResponseData, "The FolderSync element should not be null.");

            #region Capture code for FolderSync
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3096");

            // If the schema validation result is true and FolderSync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3096,
                @"[In FolderSync] The FolderSync element is a required element in FolderSync command requests and FolderSync command responses that identifies the body of the HTTP POST as containing a FolderSync command (section 2.2.2.4).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1715");

            // If the schema validation result is true and FolderSync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1715,
                @"[In FolderSync] None [Element FolderSync in FolderSync command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1716");

            // If the schema validation result is true and FolderSync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1716,
                @"[In FolderSync] Element FolderSync in FolderSync command response, the child elements are SyncKey, Status, (section 2.2.3.164.4), Changes (section 2.2.3.25).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1717");

            // If the schema validation result is true and FolderSync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1717,
                @"[In FolderSync] Element FolderSync in FolderSync command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1718");

            // If the schema validation result is true and FolderSync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1718,
                @"[In FolderSync] Element FolderSync in FolderSync command response, the number allowed is  1…1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for SyncKey(FolderSync)
            if (folderSyncResponse.ResponseData.SyncKey != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2827");

                // If the schema validation result is true and SyncKey(FolderSync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2827,
                    @"[In SyncKey(FolderSync)] Element SyncKey in FolderSync command response, the parent element is FolderSync.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2828");

                // If the schema validation result is true and SyncKey(FolderSync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2828,
                    @"[In SyncKey(FolderSync)] None [Element SyncKey in FolderSync command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2829");

                // If the schema validation result is true and SyncKey(FolderSync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2829,
                    @"[In SyncKey(FolderSync)] Element SyncKey in FolderSync command response, the child element is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2830");

                // If the schema validation result is true and SyncKey(FolderSync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2830,
                    @"[In SyncKey(FolderSync)] Element SyncKey in FolderSync command response, the number allowed is 0…1 (optional).");

                this.VerifyStringDataType();
            }
            #endregion

            #region Capture code for Status(FolderSync)
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4066");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4066,
                @"[In Status(FolderSync)] The Status element is a required child element of the FolderSync element in FolderSync command responses that indicates the success or failure of a FolderSync command request (section 2.2.2.4).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2703");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2703,
                @"[In Status(FolderSync)] Element Status in FolderSync command response, the parent element is FolderSync (section 2.2.3.71).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2704");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2704,
                @"[In Status(FolderSync)] None [Element Status in FolderSync command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2705");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2705,
                @"[In Status(FolderSync)] Element Status in FolderSync command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2706");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2706,
                @"[In Status(FolderSync)] Element Status in FolderSync command response, the number allowed is 1…1 (required).");

            Common.VerifyActualValues("Status(FolderSync)", AdapterHelper.ValidStatus(new string[] { "1", "6", "9", "10", "11", "12" }), folderSyncResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4071");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4071
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4071,
                @"[In Status(FolderSync)] The following table lists the status codes [1,6,9,10,11,12] for the FolderSync command (section 2.2.2.4). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4094");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4094,
                @"[In Status(FolderUpdate)] The Status element is a required child element of the FolderUpdate element in FolderUpdate command responses that indicates the success or failure of a FolderUpdate command request (section 2.2.2.5).");

            this.VerifyIntegerDataType();

            #endregion

            #region Capture code for Changes
            if (folderSyncResponse.ResponseData.Changes != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1155");

                // If the schema validation result is true and Changes is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1155,
                    @"[In Changes] Element Changes in FolderSync command response (section 2.2.2.4), the parent element is FolderSync (section 2.2.3.71).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1156");

                // If the schema validation result is true and Changes is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1156,
                    @"[In Changes] Element Changes in FolderSync command response (section 2.2.2.4), the child elements are Count (section 2.2.3.37), Update (section 2.2.3.177), Delete (section 2.2.3.42.1), Add (section 2.2.3.7.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1157");

                // If the schema validation result is true and Changes is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1157,
                    @"[In Changes]  Element Changes in FolderSync command response (section 2.2.2.4), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1158");

                // If the schema validation result is true and Changes is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1158,
                    @"[In Changes] Element Changes in FolderSync command response (section 2.2.2.4), the number allowed is  0…1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Count
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1454");

                // If the schema validation result is true and Count is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1454,
                    @"[In Count] Element Count in FolderSync command response (section 2.2.2.4), the parent element is Changes (section 2.2.3.25).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1455");

                // If the schema validation result is true and Count is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1455,
                    @"[In Count] None [Element Count in FolderSync command response (section 2.2.2.4) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1456");

                // Verify MS-ASCMD requirement: MS-ASCMD_R1456
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    folderSyncResponse.ResponseData.Changes.Count.GetType(),
                    1456,
                    @"[In Count] Element Count in FolderSync command response (section 2.2.2.4), the data type is unsigned integer ([MS-ASDTYPE] section 2.6).");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1457");

                // If the schema validation result is true and Count is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1457,
                    @"[In Count] Element Count in FolderSync command response (section 2.2.2.4), the number allowed is 0…1 (optional).");

                #endregion

                #region Capture code for Delete(FolderSync)
                if (folderSyncResponse.ResponseData.Changes.Delete != null && folderSyncResponse.ResponseData.Changes.Delete.Length > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1482");

                    // If the schema validation result is true and Delete(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1482,
                        @"[In Delete(FolderSync)] Element Delete in FolderSync command response (section 2.2.2.4), the parent element is Changes (section 2.2.3.25).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1483");

                    // If the schema validation result is true and Delete(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1483,
                        @"[In Delete(FolderSync)] Element Delete in FolderSync command response (section 2.2.2.4), the child element is ServerId (section 2.2.3.156.3).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1484");

                    // If the schema validation result is true and Delete(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1484,
                        @"[In Delete(FolderSync)] Element Delete in FolderSync command response (section 2.2.2.4), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1485");

                    // If the schema validation result is true and Delete(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1485,
                        @"[In Delete(FolderSync)] Element Delete in FolderSync command response (section 2.2.2.4), the number allowed is 0...N (optional).");

                    this.VerifyContainerDataType();

                    foreach (FolderSyncChangesDelete folderSyncChangesDelete in folderSyncResponse.ResponseData.Changes.Delete)
                    {
                        #region Capture code for ServerId
                        Site.Assert.IsNotNull(folderSyncChangesDelete.ServerId, "The ServerId(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3911");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3911,
                            @"[In ServerId(FolderSync)] The ServerId element is a required child element of the Update element, the Delete element, and the Add element in FolderSync command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3916");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3916,
                            @"[In ServerId(FolderSync)] Each Update element, each Delete element, and each Add element included in a FolderSync response MUST contain one ServerId element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2581");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2581,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response, the parent element is Delete (section 2.2.3.42.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2582");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                           2582,
                           @"[In ServerId(FolderSync)] None [Element ServerId in FolderSync command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2583");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                           2583,
                           @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response,  the data type is string.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2584");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2584,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response, the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();
                        #endregion
                    }
                }
                #endregion

                #region Capture code for Update(FolderSync)
                if (folderSyncResponse.ResponseData.Changes.Update != null && folderSyncResponse.ResponseData.Changes.Update.Length > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2879");

                    // If the schema validation result is true and Update is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2879,
                        @"[In Update] Element Update in FolderSync command response (section 2.2.2.4), the parent element is Changes (section 2.2.3.25).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2880");

                    // If the schema validation result is true and Update is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2880,
                        @"[In Update] Element Update in FolderSync command response (section 2.2.2.4), the child elements are ServerId (section 2.2.3.156.3), ParentId (section 2.2.3.123.2), DisplayName (section 2.2.3.47.3), Type (section 2.2.3.176.3).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2881");

                    // If the schema validation result is true and Update is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2881,
                        @"[In Update] Element Update in FolderSync command response (section 2.2.2.4), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2882");

                    // If the schema validation result is true and Update is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2882,
                        @"[In Update] Element Update in FolderSync command response (section 2.2.2.4), the number allowed is 0...N (optional).");

                    this.VerifyContainerDataType();

                    foreach (FolderSyncChangesUpdate folderSyncChangesUpdate in folderSyncResponse.ResponseData.Changes.Update)
                    {
                        #region Capture code for ServerId
                        Site.Assert.IsNotNull(folderSyncChangesUpdate.ServerId, "The ServerId(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3911");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3911,
                            @"[In ServerId(FolderSync)] The ServerId element is a required child element of the Update element, the Delete element, and the Add element in FolderSync command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3916");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3916,
                            @"[In ServerId(FolderSync)] Each Update element, each Delete element, and each Add element included in a FolderSync response MUST contain one ServerId element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2577");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                           2577,
                           @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response (section 2.2.2.4), the parent element is Update (section 2.2.3.177).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2578");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2578,
                            @"[In ServerId(FolderSync)] None [Element ServerId in FolderSync command response (section 2.2.2.4) has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2579");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2579,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response (section 2.2.2.4), the data type is string ([MS-ASDTYPE] section 2.7).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2580");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                             2580,
                             @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response (section 2.2.2.4), the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();
                        #endregion

                        #region Capture code for ParentId(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesUpdate.ParentId, "The ParentId(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3622");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3622,
                            @"[In ParentId(FolderSync)] The ParentId element is a required child element of the Update element in FolderSync command responses that specifies the server ID of the parent folder of the folder on the server that has been updated.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3625");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3625,
                            @"[In ParentId(FolderSync)] Each Update element included in a FolderSync response MUST contain one ParentId element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2332");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2332,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response (section 2.2.2.4), the parent element is Update (section 2.2.3.177).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2333");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2333,
                            @"[In ParentId(FolderSync)] None [Element ParentId in FolderSync command response (section 2.2.2.4) has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2334");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2334,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response (section 2.2.2.4), the data type is string ([MS-ASDTYPE] section 2.7).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2335");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2335,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response (section 2.2.2.4), the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();

                        #endregion

                        #region Capture code for DisplayName(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesUpdate.DisplayName, "The DisplayName(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2195");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2195,
                            @"[In DisplayName(FolderSync)] The DisplayName element is a required child element of the Update element and the Add element in FolderSync command responses that specifies the name of the folder that is shown to the user.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1530");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1530,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response (section 2.2.2.4), the parent element is Update (section 2.2.3.177).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1531");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1531,
                            @"[In DisplayName(FolderSync)] None [Element DisplayName in FolderSync command response (section 2.2.2.4) has no child element .]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1532");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1532,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response (section 2.2.2.4), the data  type is string ([MS-ASDTYPE] section 2.7).");

                        this.VerifyStringDataType();

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1533");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1533,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response (section 2.2.2.4), the number allowed is 1…1 (required).");

                        #endregion

                        #region Capture code for Type(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesUpdate.Type, "The Type(FolderSync) should not be null.");

                        int type;

                        Site.Assert.IsTrue(int.TryParse(folderSyncChangesUpdate.Type, out type), "The Type element should be an integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4663");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            4663,
                            @"[In Type(FolderSync)] The Type element is a required child element of the Update element and the Add element in FolderSync command responses that specifies the type of the folder that was updated (renamed or moved) or added on the server.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4665");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            4665,
                            @"[In Type(FolderSync)] Each Update element and each Add element included in a FolderSync response MUST contain one Type element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2867");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2867,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response (section 2.2.2.4), the parent element is Update (section 2.2.3.177).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2868");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2868,
                            @"[In Type(FolderSync)] None [Element Type in FolderSync command response (section 2.2.2.4) has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2870");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2870,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response (section 2.2.2.4), the number allowed is 1…1 (required).");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2869
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2869");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2869
                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2869,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response (section 2.2.2.4), the data type is integer ([MS-ASDTYPE] section 2.6).");

                        Common.VerifyActualValues("Type(FolderSync)", new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19" }, folderSyncChangesUpdate.Type, this.Site);

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4666");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4666
                        // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                        Site.CaptureRequirement(
                            4666,
                            @"[In Type(FolderSync)] The folder type values [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19] are listed in the following table.");

                        this.VerifyIntegerDataType();
                        #endregion
                    }
                }
                #endregion

                #region Capture code for Add(FolderSync)
                if (folderSyncResponse.ResponseData.Changes.Add != null && folderSyncResponse.ResponseData.Changes.Add.Length > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1031");

                    // If the schema validation result is true and Add(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1031,
                        @"[In Add(FolderSync)] Element Add in FolderSync command response (section 2.2.2.4), the parent element is Changes (section 2.2.3.25).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1032");

                    // If the schema validation result is true and Add(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1032,
                        @"[In Add(FolderSync)] Element Add in FolderSync command response (section 2.2.2.4), the child elements are ServerId (section 2.2.3.156.3), ParentId (section 2.2.3.123.2), DisplayName (section 2.2.3.47.3), Type (section 2.2.3.176.3).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1033");

                    // If the schema validation result is true and Add(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1033,
                        @"[In Add(FolderSync)] Element Add in FolderSync command response (section 2.2.2.4), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1034");

                    // If the schema validation result is true and Add(FolderSync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1034,
                        @"[In Add(FolderSync)] Element Add in FolderSync command response (section 2.2.2.4), the number allowed is 0...N (optional).");

                    this.VerifyContainerDataType();

                    foreach (FolderSyncChangesAdd folderSyncChangesAdd in folderSyncResponse.ResponseData.Changes.Add)
                    {
                        #region Capture code for ServerId(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesAdd.ServerId, "The ServerId(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3911");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3911,
                            @"[In ServerId(FolderSync)] The ServerId element is a required child element of the Update element, the Delete element, and the Add element in FolderSync command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3916");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            3916,
                            @"[In ServerId(FolderSync)] Each Update element, each Delete element, and each Add element included in a FolderSync response MUST contain one ServerId element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2585");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2585,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response, the parent element is Add (section 2.2.3.7.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2586");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2586,
                            @"[In ServerId(FolderSync)] None [Element ServerId in FolderSync command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2587");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2587,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response, the  data type is string.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2588");

                        // If the schema validation result is true and ServerId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2588,
                            @"[In ServerId(FolderSync)] Element ServerId in FolderSync command response, the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();
                        #endregion

                        #region Capture code for ParentId(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesAdd.ParentId, "The ParentId(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5884");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5884,
                            @"[In ParentId(FolderSync)] The ParentId element is a required child element of the Add element in FolderSync command responses that specifies the server ID of the parent folder of the folder on the server that has been added.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5885");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5885,
                            @"[In ParentId(FolderSync)] Each Add element included in a FolderSync response MUST contain one ParentId element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2336");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2336,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response, the parent element is Add (section 2.2.3.7.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2337");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2337,
                            @"[In ParentId(FolderSync)] None [Element ParentId in FolderSync command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2338");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2338,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response, the data type is string.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2339");

                        // If the schema validation result is true and ParentId(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2339,
                            @"[In ParentId(FolderSync)] Element ParentId in FolderSync command response, the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();

                        #endregion

                        #region Capture code for DisplayName(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesAdd.DisplayName, "The DisplayName(FolderSync) should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2195");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2195,
                            @"[In DisplayName(FolderSync)] The DisplayName element is a required child element of the Update element and the Add element in FolderSync command responses that specifies the name of the folder that is shown to the user.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1539");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1539,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response, the parent element is Add (section 2.2.3.7.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1540");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1540,
                            @"[In DisplayName(FolderSync)] None [Element DisplayName in FolderSync command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1541");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1541,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response, the data type is string.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1542");

                        // If the schema validation result is true and DisplayName(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1542,
                            @"[In DisplayName(FolderSync)] Element DisplayName in FolderSync command response, the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();
                        #endregion

                        #region Capture code for Type(FolderSync)
                        Site.Assert.IsNotNull(folderSyncChangesAdd.Type, "The Type(FolderSync) should not be null.");

                        int type;

                        Site.Assert.IsTrue(int.TryParse(folderSyncChangesAdd.Type, out type), "The Type element should be an integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4663");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            4663,
                            @"[In Type(FolderSync)] The Type element is a required child element of the Update element and the Add element in FolderSync command responses that specifies the type of the folder that was updated (renamed or moved) or added on the server.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4665");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            4665,
                            @"[In Type(FolderSync)] Each Update element and each Add element included in a FolderSync response MUST contain one Type element.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2871");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2871,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response, the parent element is Add (section 2.2.3.7.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2872");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2872,
                            @"[In Type(FolderSync)] None [Element Type in FolderSync command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2874");

                        // If the schema validation result is true and Type(FolderSync) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2874,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response, the number allowed is 1…1 (required).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2873");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2873
                        Site.CaptureRequirement(
                            2873,
                            @"[In Type(FolderSync)] Element Type in FolderSync command response, the data type is integer.");

                        Common.VerifyActualValues("Type(FolderSync)", new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19" }, folderSyncChangesAdd.Type, this.Site);

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4666");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4666
                        // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                        Site.CaptureRequirement(
                          4666,
                            @"[In Type(FolderSync)] The folder type values [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19] are listed in the following table.");

                        this.VerifyIntegerDataType();

                        #endregion
                    }
                }
                #endregion
            }
            #endregion
        }

        #endregion

        #region Capture code for FolderUpdate command
        /// <summary>
        /// This method is used to verify the FolderUpdate response related requirements.
        /// </summary>
        /// <param name="folderUpdateResponse">FolderUpdate command response.</param>
        private void VerifyFolderUpdateCommand(FolderUpdateResponse folderUpdateResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(folderUpdateResponse.ResponseData, "The FolderUpdate element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3099");

            // If the schema validation result is true and FolderUpdate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3099,
                @"[In FolderUpdate] The FolderUpdate element is a required element in FolderUpdate command requests and FolderUpdate command responses that identifies the body of the HTTP POST as containing a FolderUpdate command (section 2.2.2.5).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1723");

            // If the schema validation result is true and FolderUpdate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1723,
                @"[In FolderUpdate] None [Element FolderUpdate in FolderUpdate command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1724");

            // If the schema validation result is true and FolderUpdate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1724,
                @"[In FolderUpdate] Element FolderUpdate in FolderUpdate command response, the child elements are SyncKey, Status (section 2.2.3.162.5).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1725");

            // If the schema validation result is true and FolderUpdate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1725,
                @"[In FolderUpdate] Element FolderUpdate in FolderUpdate command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1726");

            // If the schema validation result is true and FolderUpdate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1726,
                @"[In FolderUpdate] Element FolderUpdate in FolderUpdate command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #region Capture code for SyncKey(FolderUpdate)
            if (folderUpdateResponse.ResponseData.SyncKey != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2819");

                // If the schema validation result is true and SyncKey(FolderUpdate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2819,
                    @"[In SyncKey(FolderUpdate)] Element SyncKey in FolderUpdate command response, the parent element is FolderUpdate.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2820");

                // If the schema validation result is true and SyncKey(FolderUpdate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2820,
                    @"[In SyncKey(FolderUpdate)] None [Element SyncKey in FolderUpdate command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2821");

                // If the schema validation result is true and SyncKey(FolderUpdate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2821,
                    @"[In SyncKey(FolderUpdate)] Element SyncKey in FolderUpdate command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2822");

                // If the schema validation result is true and SyncKey(FolderUpdate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2822,
                    @"[In SyncKey(FolderUpdate)] Element SyncKey in FolderUpdate command response, the number allowed is 0…1 (optional).");

                this.VerifyStringDataType();
            }
            #endregion

            #region Capture code for Status(FolderUpdate)

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2707");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2707,
                @"[In Status(FolderUpdate)] Element Status in FolderUpdate command response,the parent element is FolderUpdate (section 2.2.3.72).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2708");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2708,
                @"[In Status(FolderUpdate)] None [Element Status in FolderUpdate command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2710");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2710,
                @"[In Status(FolderUpdate)] Element Status in FolderUpdate command response, the number allowed is 1…1 (required).");

            Common.VerifyActualValues("Status(FolderUpdate)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "9", "10", "11" }), folderUpdateResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4097");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4097
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4097,
                @"[In Status(FolderUpdate)] The following table lists the status codes [1,2,3,4,5,6,9,10,11] for the FolderUpdate command (section 2.2.2.5). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

            this.VerifyIntegerDataType();

            #endregion
        }

        #endregion

        #region Capture code for GetItemEstimate command
        /// <summary>
        /// This method is used to verify the GetItemEstimate response related requirements.
        /// </summary>
        /// <param name="getItemEstimateResponse">GetItemEstimate command response.</param>
        private void VerifyGetItemEstimateCommand(Microsoft.Protocols.TestSuites.Common.GetItemEstimateResponse getItemEstimateResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(getItemEstimateResponse.ResponseData, "The GetItemEstimate element should not be null.");

            #region Capture code for GetItemEstimate
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3130");

            // If the schema validation result is true and GetItemEstimate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3130,
                @"[In GetItemEstimate] The GetItemEstimate element is a required element in GetItemEstimate command requests and GetItemEstimate command responses that identifies the body of the HTTP POST as containing a GetItemEstimate command (section 2.2.2.7).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1768");

            // If the schema validation result is true and GetItemEstimate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1768,
                @"[In GetItemEstimate] None [Element GetItemEstimate in GetItemEstimate command response has no parent element .]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1769");

            // If the schema validation result is true and GetItemEstimate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1769,
                @"[In GetItemEstimate] Element GetItemEstimate in GetItemEstimate command response, the child element is Response (section 2.2.3.144.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1770");

            // If the schema validation result is true and GetItemEstimate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1770,
                @"[In GetItemEstimate] Element GetItemEstimate in GetItemEstimate command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1771");

            // If the schema validation result is true and GetItemEstimate is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1771,
                @"[In GetItemEstimate] Element GetItemEstimate in GetItemEstimate command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Status
            if (!string.IsNullOrEmpty(getItemEstimateResponse.ResponseData.Status))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2711");

                // If the schema validation result is true and Status(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2711,
                    @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the parent element is GetItemEstimate (section 2.2.3.81).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2712");

                // If the schema validation result is true and Status(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2712,
                    @"[In Status(GetItemEstimate)] None [Element Status in GetItemEstimate command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2713");

                // If the schema validation result is true and Status(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2713,
                    @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2714");

                // If the schema validation result is true and Status(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2714,
                    @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the number allowed is 0…1 (optional).");

                Common.VerifyActualValues("Status(GetItemEstimate)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4" }), getItemEstimateResponse.ResponseData.Status.ToString(), this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4131");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4131
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4131,
                    @"[In Status(GetItemEstimate)] The following table lists the status codes [1,2,3,4] for the GetItemEstimate command (section 2.2.2.8). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

                this.VerifyIntegerDataType();
            }
            #endregion

            #region Capture code for Response(GetItemEstimate)
            if (getItemEstimateResponse.ResponseData.Response != null && getItemEstimateResponse.ResponseData.Response.Length > 0)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3859");

                // If the schema validation result is true and Response(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    3859,
                    @"[In Response(GetItemEstimate)] The Response element is a required child element of the GetItemEstimate element in GetItemEstimate command responses that contains elements that describe estimated changes.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2483");

                // If the schema validation result is true and Response(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2483,
                    @"[In Response(GetItemEstimate)] Element Response in GetItemEstimate command response (section 2.2.2.8), the parent element is GetItemEstimate (section 2.2.3.81).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2484");

                // If the schema validation result is true and Response(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2484,
                    @"[In Response(GetItemEstimate)] Element Response in GetItemEstimate command response (section 2.2.2.8), the child elements are Status (section 2.2.3.167.6), Collection (section 2.2.3.29.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2485");

                // If the schema validation result is true and Response(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2485,
                    @"[In Response(GetItemEstimate)] Element Response in GetItemEstimate command response (section 2.2.2.8), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2486");

                // If the schema validation result is true and Response(GetItemEstimate) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2486,
                    @"[In Response(GetItemEstimate)] Element Response in GetItemEstimate command response (section 2.2.2.8), the number allowed is 1…N (required).");

                this.VerifyContainerDataType();
            }
            #endregion

            if (getItemEstimateResponse.ResponseData.Response != null)
            {
                foreach (TestSuites.Common.Response.GetItemEstimateResponse response in getItemEstimateResponse.ResponseData.Response)
                {
                    #region Capture code for Status
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4127");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        4127,
                        @"[In Status(GetItemEstimate)] The Status element is a required child element of the Response element in GetItemEstimate command responses that indicates the success or failure of part or all of a GetItemEstimate command request (section 2.2.2.8).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5577");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        5577,
                        @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the parent element is Response (section 2.2.3.144.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5578");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        5578,
                        @"[In Status(GetItemEstimate)] None [Element Status in GetItemEstimate command response has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5579");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        5579,
                        @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the data type is unsignedByte.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5580");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        5580,
                        @"[In Status(GetItemEstimate)] Element Status in GetItemEstimate command response, the number allowed is 1…1 (required).");

                    if (!string.IsNullOrEmpty(getItemEstimateResponse.ResponseData.Status))
                    {
                        Common.VerifyActualValues("Status(GetItemEstimate)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4" }), getItemEstimateResponse.ResponseData.Status.ToString(), this.Site);

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4131");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4131
                        // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                        Site.CaptureRequirement(
                            4131,
                            @"[In Status(GetItemEstimate)] The following table lists the status codes [1,2,3,4] for the GetItemEstimate command (section 2.2.2.7). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");
                    }

                    this.VerifyIntegerDataType();
                    #endregion

                    #region Capture code for Collection
                    if (response.Collection != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1355");

                        // If the schema validation result is true and Collection(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1355,
                            @"[In Collection(GetItemEstimate)] Element Collection in GetItemEstimate command response , the parent element is Response (section 2.2.3.144.2).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1356");

                        // If the schema validation result is true and Collection(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1356,
                            @"[In Collection(GetItemEstimate)] Element Collection in GetItemEstimate command response, the child elements are Class (section 2.2.27.1), CollectionId (section 2.2.3.30.1), Estimate (section 2.2.3.62).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1357");

                        // If the schema validation result is true and Collection(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1357,
                            @"[In Collection(GetItemEstimate)] Element Collection in GetItemEstimate command response, the data type is container.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1358");

                        // If the schema validation result is true and Collection(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1358,
                            @"[In Collection(GetItemEstimate)] Element Collection in GetItemEstimate command response, the number allowed is  0…1 (optional).");

                        this.VerifyContainerDataType();

                        #region Capture code for CollectionId
                        Site.Assert.IsNotNull(response.Collection.CollectionId, "The CollectionId should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R983");

                        // If the schema validation result is true and CollectionId(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            983,
                            @"[In CollectionId(GetItemEstimate)] The CollectionId element is a required child element of the Collection element in GetItemEstimate command requests and responses that specifies the server ID of the collection from which the item estimate is being obtained.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1371");

                        // If the schema validation result is true and CollectionId(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1371,
                            @"[In CollectionId(GetItemEstimate)] Element CollectionId in GetItemEstimate command response, the parent element is Collection.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1372");

                        // If the schema validation result is true and CollectionId(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1372,
                            @"[In CollectionId(GetItemEstimate)] None [Element CollectionId in GetItemEstimate command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1374");

                        // If the schema validation result is true and CollectionId(GetItemEstimate) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1374,
                            @"[In CollectionId(GetItemEstimate)] Element CollectionId in GetItemEstimate command response (section 2.2.2.7.2), the number allowed is 1...1 (required).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1373");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R1373
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(string),
                            response.Collection.CollectionId.GetType(),
                            1373,
                            @"[In CollectionId(GetItemEstimate)] Element CollectionId in GetItemEstimate command response, the data type is string.");

                        this.VerifyStringDataType();

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5869");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R5869
                        Site.CaptureRequirementIfIsTrue(
                            response.Collection.CollectionId.Length <= 64,
                            5869,
                            @"[In CollectionId(GetItemEstimate)] The CollectionId element value is not larger than 64 characters in length.");
                        #endregion

                        #region Capture code for Estimate
                        if (response.Collection.Estimate != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2289");

                            // If the schema validation result is true and Estimate is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2289,
                                @"[In Estimate] The Estimate element is a required child element of the Collection element in GetItemEstimate command responses that specifies the estimated number of items in the collection or folder that have to be synchronized.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1643");

                            // If the schema validation result is true and Estimate is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1643,
                                @"[In Estimate] Element Estimate in GetItemEstimate command response (section 2.2.2.7), the parent element is Collection (section 2.2.3.29.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1644");

                            // If the schema validation result is true and Estimate is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1644,
                                @"[In Estimate] None [Element Estimate in GetItemEstimate command response (section 2.2.2.7) has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1646");

                            // If the schema validation result is true and Estimate is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1646,
                                @"[In Estimate] Element Estimate in GetItemEstimate command response (section 2.2.2.7), the number allowed is 1…1 (required).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1645");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R1645
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(int),
                                Convert.ToInt32(response.Collection.Estimate).GetType(),
                                1645,
                                @"[In Estimate] Element Estimate in GetItemEstimate command response (section 2.2.2.7), the data type is integer ([MS-ASDTYPE] section 2.5).");

                            this.VerifyIntegerDataType();
                        }
                        #endregion
                    }
                    #endregion
                }
            }
        }
        #endregion

        #region Capture code for ItemOperations command
        /// <summary>
        /// This method is used to verify the ItemOperations response related requirements.
        /// </summary>
        /// <param name="itemOperationsResponse">ItemOperations command response.</param>
        private void VerifyItemOperationsCommand(Microsoft.Protocols.TestSuites.Common.ItemOperationsResponse itemOperationsResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(itemOperationsResponse.ResponseData, "The ItemOperations element should not be null.");

            #region Capture code for MultiPartResponse
            if (itemOperationsResponse.Headers["Content-Type"].Split(new char[] { ';' })[0].Equals("application/vnd.ms-sync.multipart"))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5506");

                // If the schema validation result is true, ItemOperations is not null and the content is in multipart, this requirement can be verified.
                Site.CaptureRequirement(
                    5506,
                    @"[In Delivery of Content Requested by Fetch] The format of the body of the response is a MultiPartResponse structure, specified in section 2.2.2.9.1.1.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5511");

                // If the schema validation result is true, ItemOperations is not null and the content is in multipart, this requirement can be verified.
                Site.CaptureRequirement(
                    5511,
                    @"[In MultiPartResponse] PartsMetaData (variable): This field [PartsMetaData] is an array of PartMetaData structures, as specified in section 2.2.2.9.1.1.1.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5513");

                // If the schema validation result is true, ItemOperations is not null and the content is in multipart, this requirement can be verified.
                Site.CaptureRequirement(
                    5513,
                    @"[In MultiPartResponse] Parts (variable): This field [Parts] is an array of bytes that contains the data for the parts of the multipart response.");
            }
            #endregion

            #region Capture code for ItemOperations
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3364");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3364,
                @"[In MIMESupport(ItemOperations)] The airsyncbase:Body element is a complex element");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3206");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3206,
                @"[In ItemOperations] The ItemOperations element is a required element in ItemOperations command requests and ItemOperations command responses that identifies the body of the HTTP POST as containing an ItemOperations command (section 2.2.2.8).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1820");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1820,
                @"[In ItemOperations] None [Element ItemOperations in ItemOperations command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1821");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1821,
                @"[In ItemOperations] Element ItemOperations in ItemOperations command response, the child elements are Status (section 2.2.3.167.7), Response (section 2.2.3.144.3).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1822");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1822,
                @"[In ItemOperations] Element ItemOperations in ItemOperations command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1823");

            // If the schema validation result is true and ItemOperations is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1823,
                @"[In ItemOperations] Element ItemOperations in ItemOperations command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Status
            int status;

            Site.Assert.IsTrue(int.TryParse(itemOperationsResponse.ResponseData.Status, out status), "The Status element should be an integer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R194");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                194,
                @"[In ItemOperations] The server MUST report the status per operation to the client.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2715");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2715,
                @"[In Status(ItemOperations)] Element Status in ItemOperations command response, the parent element is ItemOperations (section 2.2.3.89).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2716");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2716,
                @"[In Status(ItemOperations)] None [Element Status in ItemOperations command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2718");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2718,
                @"[In Status(ItemOperations)] Element Status in ItemOperations command response, the number allowed is 1...1  (required).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2717");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2717
            Site.CaptureRequirement(
                2717,
                @"[In Status(ItemOperations)] Element Status in ItemOperations command response, the data type is integer ([MS-ASDTYPE] section 2.6).");

            Common.VerifyActualValues("Status(ItemOperations)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "14", "15", "16", "17", "18", "155", "156" }), itemOperationsResponse.ResponseData.Status, this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4151");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4151
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4151,
                @"[In Status(ItemOperations)] The following table lists the status codes [1,2,3,4,5,6,7,8,9,10,11,12,14,15,16,17,18,155,156] for the ItemOperations command (section 2.2.2.9). For information about status values common to all ActiveSync commands, see section 2.2.4.");

            this.VerifyIntegerDataType();
            #endregion

            #region Capture code for Response
            if (itemOperationsResponse.ResponseData.Response != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2487");

                // If the schema validation result is true and Response(ItemOperations) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2487,
                    @"[In Response(ItemOperations)] Element Response in ItemOperations command response (section 2.2.2.9), the parent element is ItemOperations.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2488");

                // If the schema validation result is true and Response(ItemOperations) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2488,
                    @"[In Response(ItemOperations)] Element Response in ItemOperations command response (section 2.2.2.9), the child elements are EmptyFolderContents (section 2.2.3.55), Fetch (section 2.2.3.63.1), Move (section 2.2.3.111.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2489");

                // If the schema validation result is true and Response(ItemOperations) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2489,
                    @"[In Response(ItemOperations)] Element Response in ItemOperations command response (section 2.2.2.9), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2490");

                // If the schema validation result is true and Response(ItemOperations) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2490,
                    @"[In Response(ItemOperations)] Element Response in ItemOperations command response (section 2.2.2.9), the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for EmptyFolderContents
                if (itemOperationsResponse.ResponseData.Response.EmptyFolderContents != null && itemOperationsResponse.ResponseData.Response.EmptyFolderContents.Length > 0)
                {
                    foreach (ItemOperationsResponseEmptyFolderContents content in itemOperationsResponse.ResponseData.Response.EmptyFolderContents)
                    {
                        Site.Assert.IsNotNull(content, "The EmptyFolderContent should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1603");

                        // If the schema validation result is true and EmptyFolderContents is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1603,
                            @"[In EmptyFolderContents] Element EmptyFolderContents in ItemOperations command response, the parent element is Response (section 2.2.3.140.3).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1604");

                        // If the schema validation result is true and EmptyFolderContents is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1604,
                            @"[In EmptyFolderContents] Element EmptyFolderContents in ItemOperations command response, the child elements are airsync:CollectionId (section 2.2.3.30.2)
,Status (section 2.2.3.162.7).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1605");

                        // If the schema validation result is true and EmptyFolderContents is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1605,
                            @"[In EmptyFolderContents] Element EmptyFolderContents in ItemOperations command response, the data type is container.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1606");

                        // If the schema validation result is true and EmptyFolderContents is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1606,
                            @"[In EmptyFolderContents] Element EmptyFolderContents in ItemOperations command response, the number allowed is 0...N (optional).");

                        this.VerifyContainerDataType();

                        #region Capture code for Status
                        int statusForEmptyFolderContents;

                        Site.Assert.IsTrue(int.TryParse(content.Status, out statusForEmptyFolderContents), "As a child element of EmptyFolderContents, the Status element should be an integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4148");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4148
                        // If the above assert passed, this requirement can be verified.
                        Site.CaptureRequirement(
                            4148,
                            @"[In Status(ItemOperations)] The Status element is a required child element of the ItemOperations element, the Move element, the EmptyFolderContents element, and the Fetch element in ItemOperations command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2723");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2723
                        // If the above assert passed, this requirement can be verified.
                        Site.CaptureRequirement(
                            2723,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response EmptyFolderContents operation, the parent element is EmptyFolderContents (section 2.2.3.55).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2724");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2724,
                            @"[In Status(ItemOperations)] None [Element Status in ItemOperations command response EmptyFolderContents operation has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2726");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2726,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response EmptyFolderContents operation, the number allowed is 1...1 (required).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2725");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2725,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response EmptyFolderContents operation, the data type is integer.");

                        this.VerifyIntegerDataType();
                        #endregion

                        #region Capture code for CollectionId
                        Site.Assert.IsNotNull(content.CollectionId, "CollectionId in EmptyFolderContents should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1379");

                        // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1379,
                            @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the parent elements are EmptyFolderContents, Fetch.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1380");

                        // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1380,
                            @"[In CollectionId(ItemOperations)] None [Element CollectionId in ItemOperations command response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1381");

                        // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1381,
                            @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the data type is string.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5414");

                        // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5414,
                            @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the number allowed is 1…1 (required in EmptyFolderContents operation).");

                        this.VerifyStringDataType();
                        #endregion
                    }
                }

                #endregion

                #region Capture code for Move

                if (itemOperationsResponse.ResponseData.Response.Move != null && itemOperationsResponse.ResponseData.Response.Move.Length > 0)
                {
                    foreach (ItemOperationsResponseMove move in itemOperationsResponse.ResponseData.Response.Move)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1959");

                        // If the schema validation result is true and Move(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1959,
                            @"[In Move(ItemOperations)] Element Move in ItemOperations command response, the parent element is Response (section 2.2.3.144.3).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1960");

                        // If the schema validation result is true and Move(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1960,
                            @"[In Move(ItemOperations)] Element Move in ItemOperations command response, the child elements are ConversationId,Status (section 2.2.3.167.7).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1961");

                        // If the schema validation result is true and Move(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1961,
                            @"[In Move(ItemOperations)] Element Move in ItemOperations command response, the data type is container.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1962");

                        // If the schema validation result is true and Move(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1962,
                            @"[In Move(ItemOperations)] Element Move in ItemOperations command response, the number allowed is 0...N (optional).");

                        this.VerifyContainerDataType();

                        #region Capture code for Status
                        int statusForMove;

                        Site.Assert.IsTrue(int.TryParse(move.Status, out statusForMove), "As a child element of Move, the Status element should be an integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4148");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4148
                        // If the above assert passed, this requirement can be verified.
                        Site.CaptureRequirement(
                            4148,
                            @"[In Status(ItemOperations)] The Status element is a required child element of the ItemOperations element, the Move element, the EmptyFolderContents element, and the Fetch element in ItemOperations command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2719");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2719,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response Move operation, the parent element is Move (section 2.2.3.111.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2720");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2720,
                            @"[In Status(ItemOperations)] None [Element Status in ItemOperations command response Move operation has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2722");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2722,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response Move operation, the number allowed is 1...1  (required).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2721");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2721
                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2721,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response Move operation, the data type is integer.");

                        this.VerifyIntegerDataType();
                        #endregion

                        #region Capture code for ConversationId
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                        {
                            Site.Assert.IsNotNull(move.ConversationId, "The ConversationId element should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2080");

                            // If the schema validation result is true and ConversationId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2080,
                                @"[In ConversationId(ItemOperations)] The ConversationId element is a required child element of the Move element in ItemOperations command requests and responses that specifies the conversation to be moved.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1437");

                            // If the schema validation result is true and ConversationId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1437,
                                @"[In ConversationId(ItemOperations)] Element ConversationId in ItemOperations command response Move operation, the parent element is Move.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1438");

                            // If the schema validation result is true and ConversationId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1438,
                                @"[In ConversationId(ItemOperations)] None [Element ConversationId in ItemOperations command response Move operation has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1439");

                            // If the schema validation result is true and ConversationId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1439,
                                @"[In ConversationId(ItemOperations)] Element ConversationId in ItemOperations command response Move operation, the data type is string.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1440");

                            // If the schema validation result is true and ConversationId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1440,
                                @"[In ConversationId(ItemOperations)] Element ConversationId in ItemOperations command response Move operation, the number allowed is 1...1 (required).");

                            this.VerifyStringDataType();
                        }
                        #endregion
                    }
                }

                #endregion

                #region Capture code for Fetch

                if (itemOperationsResponse.ResponseData.Response.Fetch != null && itemOperationsResponse.ResponseData.Response.Fetch.Length > 0)
                {
                    foreach (ItemOperationsResponseFetch fetch in itemOperationsResponse.ResponseData.Response.Fetch)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1651");

                        // If the schema validation result is true and Fetch(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1651,
                            @"[In Fetch(ItemOperations)] Element Fetch in ItemOperations command response, the parent element is Response (section 2.2.3.140.3).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1652");

                        // If the schema validation result is true and Fetch(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1652,
                            @"[In Fetch(ItemOperations)] Element Fetch in ItemOperations command response, the child elements are documentlibrary:LinkId (optional), search:LongId (optional), airsync:CollectionId (optional), airsync:ServerId (optional), Status (section 2.2.3.162.7), airsync:Class(optional) (section 2.2.3.27.2), Properties(optional) (section 2.2.3.128.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1653");

                        // If the schema validation result is true and Fetch(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1653,
                            @"[In Fetch(ItemOperations)] Element Fetch in ItemOperations command response, the data type is container.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1654");

                        // If the schema validation result is true and Fetch(ItemOperations) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1654,
                            @"[In Fetch(ItemOperations)] Element Fetch in ItemOperations command response, the number allowed is 0...N (optional).");

                        this.VerifyContainerDataType();

                        #region Capture code for Status
                        Site.Assert.IsNotNull(fetch.Status, "Status in Fetch should not be null.");

                        int statusForFetch;

                        Site.Assert.IsTrue(int.TryParse(fetch.Status, out statusForFetch), "As a child element of Fetch, the Status element should be an integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4148");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R4148
                        // If the above assert is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            4148,
                            @"[In Status(ItemOperations)] The Status element is a required child element of the ItemOperations element, the Move element, the EmptyFolderContents element, and the Fetch element in ItemOperations command responses.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2727");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2727,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response fetch operation, the parent element is Fetch (section 2.2.3.63.1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2728");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2728,
                            @"[In Status(ItemOperations)] None [Element Status in ItemOperations command response Fetch operation has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2730");

                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2730,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response Fetch operation,  the number allowed is 1...1  (required).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2729");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2729
                        // If the schema validation result is true, this requirement can be verified.
                        Site.CaptureRequirement(
                            2729,
                            @"[In Status(ItemOperations)] Element Status in ItemOperations command response Fetch operation, the data type is integer.");

                        this.VerifyIntegerDataType();

                        #endregion

                        #region Capture code for CollectionId
                        if (fetch.CollectionId != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1379");

                            // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1379,
                                @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the parent elements are EmptyFolderContents, Fetch.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1380");

                            // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1380,
                                @"[In CollectionId(ItemOperations)] None [Element CollectionId in ItemOperations command response has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1381");

                            // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1381,
                                @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the data type is string.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5415");

                            // If the schema validation result is true and CollectionId is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                5415,
                                @"[In CollectionId(ItemOperations)] Element CollectionId in ItemOperations command response, the number allowed is 0…1 (optional in Fetch operation)");

                            this.VerifyStringDataType();
                        }

                        #endregion

                        #region Capture code for Class
                        if (fetch.Class != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1167");

                            // If the schema validation result is true and Class(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1167,
                                @"[In Class(ItemOperations)] Element Class(ItemOperations) in ItemOperations command response (section 2.2.2.9) fetch operation, the parent element is Fetch (section 2.2.3.63.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1168");

                            // If the schema validation result is true and Class(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1168,
                                @"[In Class(ItemOperations)] None [Element Class(ItemOperations) in ItemOperations command response (section 2.2.2.9) fetch operation has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1169");

                            // If the schema validation result is true and Class(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1169,
                                @"[In Class(ItemOperations)] Element Class(ItemOperations) in ItemOperations command response(section 2.2.2.9) fetch operation, the data type is string ([MS-ASDTYPE] section 2.7).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1170");

                            // If the schema validation result is true and Class(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1170,
                                @"[In Class(ItemOperations)] Element Class(ItemOperations) in ItemOperations command response (section 2.2.2.9) fetch operation, the number allowed is 0...1 (optional).");

                            Common.VerifyActualValues("Class(ItemOperations)", new string[] { "Email", "Contacts", "Calendar", "Tasks", "SMS", "Notes" }, fetch.Class, this.Site);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R913");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R913
                            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                            Site.CaptureRequirement(
                                913,
                                @"[In Class(ItemOperations)] The valid airsync:Class element values are as follows [Email, Contacts, Calendar, Tasks, SMS, Notes] for the latest protocol version.");

                            this.VerifyStringDataType();
                        }
                        #endregion

                        #region Capture code for LinkId
                        if (fetch.LinkId != null)
                        {
                            try
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1551");

                                // If the schema validation result is true, the LinkId is not null and LinkId can be converted to Uri without any exception, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1551,
                                    @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in ItemOperations command response (section 2.2.2.8) fetch operation, the parent element is itemoperations:Fetch.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1552");

                                // If the schema validation result is true, the LinkId is not null and LinkId can be converted to Uri without any exception, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1552,
                                    @"[In documentlibrary:LinkId] None [Element documentlibrary:LinkId in ItemOperations command response fetch operation has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1553");

                                // If the schema validation result is true, the LinkId is not null and LinkId can be converted to Uri without any exception, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1553,
                                    @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in ItemOperations command response fetch operation, the data type is URI.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1554");

                                // If the schema validation result is true, the LinkId is not null and LinkId can be converted to Uri without any exception, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1554,
                                    @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in ItemOperations command response fetch operation, the number allowed is 0...1 (optional).");
                            }
                            catch (UriFormatException)
                            {
                                Site.Assert.Fail(@"The LinkId ""{0}"" is not a Uri.", fetch.LinkId);
                            }
                        }
                        #endregion

                        #region Capture code for ServerId
                        if (fetch.ServerId != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2597");

                            // If the schema validation result is true and ServerId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2597,
                                @"[In ServerId(ItemOperations)] Element ServerId in ItemOperations command response Fetch operation, the parent element is Fetch.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2598");

                            // If the schema validation result is true and ServerId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2598,
                                @"[In ServerId(ItemOperations)] None [Element ServerId in ItemOperations command response Fetch operation has  no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2599");

                            // If the schema validation result is true and ServerId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2599,
                                @"[In ServerId(ItemOperations)] Element ServerId in ItemOperations command response Fetch operation, the data type is string.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2600");

                            // If the schema validation result is true and ServerId(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2600,
                                @"[In ServerId(ItemOperations)] Element ServerId in ItemOperations command response Fetch operation, the number allowed is 0...1 (optional).");

                            this.VerifyStringDataType();
                        }
                        #endregion

                        #region Capture code for Properties
                        if (fetch.Properties != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2396");

                            // If the schema validation result is true and Properties(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2396,
                                @"[In Properties(ItemOperations)] Element Properties in ItemOperations command response (section 2.2.2.9) fetch operation, the parent element is Fetch (section 2.2.3.63.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2398");

                            // If the schema validation result is true and Properties(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2398,
                                @"[In Properties(ItemOperations)] Element Properties in ItemOperations command response (section 2.2.2.9) fetch operation, the data type is container ([MS-ASDTYPE] section 2.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2399");

                            // If the schema validation result is true and Properties(ItemOperations) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2399,
                                @"[In Properties(ItemOperations)] Element Properties in ItemOperations command response (section 2.2.2.9) fetch operation, the number allowed is 0...1 (optional).");

                            this.VerifyContainerDataType();

                            for (int j = 0; j < fetch.Properties.ItemsElementName.Length; j++)
                            {
                                #region Capture code for Range
                                if (fetch.Properties.ItemsElementName[j] == ItemsChoiceType3.Range)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2415");

                                    // If the schema validation result is true and Range(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2415,
                                        @"[In Range(ItemOperations)] Element Range in ItemOperations command response Fetch operation, the parent element is Properties (section 2.2.3.132.1).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2416");

                                    // If the schema validation result is true and Range(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2416,
                                        @"[In Range(ItemOperations)] None [Element Range in ItemOperations command response Fetch operation has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2417");

                                    // If the schema validation result is true and Range(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2417,
                                        @"[In Range(ItemOperations)] Element Range in ItemOperations command response Fetch operation, the data type is string.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2418");

                                    // If the schema validation result is true and Range(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2418,
                                        @"[In Range(ItemOperations)] Element Range in ItemOperations command response Fetch operation, the number allowed is 0...1 (optional).");

                                    this.VerifyStringDataType();
                                }
                                #endregion

                                #region Capture code for Part
                                if (fetch.Properties.ItemsElementName[j] == ItemsChoiceType3.Part)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3633");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R3633
                                    Site.CaptureRequirementIfAreEqual<string>(
                                        "application/vnd.ms-sync.multipart",
                                        itemOperationsResponse.Headers["Content-Type"].ToLower(CultureInfo.InvariantCulture),
                                        3633,
                                        @"[In Part] The Part element is present only in a multipart ItemOperations response.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2344");

                                    // If the schema validation result is true and Part is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2344,
                                        @"[In Part] Element Part in ItemOperations command response (section 2.2.2.9) fetch operation, the parent elements are Properties (section 2.2.3.132.1) , airsyncbase:Body ([MS-ASAIRS] section 2.2.2.9).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2345");

                                    // If the schema validation result is true and Part is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2345,
                                        @"[In Part] None [Element Part in ItemOperations command response (section 2.2.2.9) fetch operation has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2347");

                                    // If the schema validation result is true and Part is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2347,
                                        @"[In Part] Element Part in ItemOperations command response (section 2.2.2.9) fetch operation, the number allowed is 0...1 (optional).");

                                    int part;

                                    Site.Assert.IsTrue(int.TryParse(fetch.Properties.Items[j].ToString(), out part), "The Part element should be an integer.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2346");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R2346
                                    // If the above assert is true, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2346,
                                        @"[In Part] Element Part in ItemOperations command response (section 2.2.2.9) fetch operation, the data type is integer ([MS-ASDTYPE] section 2.6).");

                                    this.VerifyIntegerDataType();
                                }
                                #endregion

                                #region Capture code for Data
                                if (fetch.Properties.ItemsElementName[j] == ItemsChoiceType3.Data)
                                {
                                    string data = (string)fetch.Properties.Items[j];

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1462");

                                    // If the schema validation result is true and Data(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1462,
                                        @"[In Data(ItemOperations)] Element Data in ItemOperations command response (section 2.2.2.9) fetch operation, the parent element is Properties (section 2.2.3.132.1).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1463");

                                    // If the schema validation result is true and Data(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1463,
                                        @"[In Data(ItemOperations)] None [Element Data in ItemOperations command response (section 2.2.2.9) fetch operation has no child element .]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1464");

                                    // If the schema validation result is true and Data(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1464,
                                        @"[In Data(ItemOperations)] Element Data in ItemOperations command response (section 2.2.2.9) fetch operation, the data type is string ([MS-ASDTYPE] section 2.7).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1465");

                                    // If the schema validation result is true and Data(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1465,
                                        @"[In Data(ItemOperations)] Element Data in ItemOperations command response (section 2.2.2.9) fetch operation, the number allowed is 0...1 (optional).");

                                    this.VerifyStringDataType();

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2130");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R2130
                                    Site.CaptureRequirementIfIsTrue(
                                        Common.IsStringBase64Encoded(data),
                                        2130,
                                        @"[In Data(ItemOperations)] The content of the Data element is a base64 encoding of the binary document, attachment, or body data.");
                                }
                                #endregion

                                #region Capture code for Version
                                if (fetch.Properties.ItemsElementName[j] == ItemsChoiceType3.Version)
                                {
                                    DateTime version = (DateTime)fetch.Properties.Items[j];

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5417");

                                    // If the schema validation result is true and Version is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        5417,
                                        @"[In Version] Element Version in ItemOperations command response (section 2.2.2.9) fetch operation, the parent element is Properties (section 2.2.3.132.1)");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5273");

                                    // If the schema validation result is true and Version is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        5273,
                                        @"[In Version] None [Element Version in ItemOperations command response (section 2.2.2.9) fetch operation has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5275");

                                    // If the schema validation result is true and Version is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        5275,
                                        @"[In Version] Element Version in ItemOperations command response (section 2.2.2.9) fetch operation, the number allowed is 0...1 (optional).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5274");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R5274
                                    Site.CaptureRequirementIfAreEqual<Type>(
                                        typeof(DateTime),
                                        version.GetType(),
                                        5274,
                                        @"[In Version] Element Version in ItemOperations command response (section 2.2.2.9) fetch operation, the data type is datetime ([MS-ASDTYPE] section 2.3).");

                                    this.VerifyDateTimeStructure();
                                }
                                #endregion

                                #region Capture code for Total
                                if (fetch.Properties.ItemsElementName[j] == ItemsChoiceType3.Total)
                                {
                                    int total;

                                    Site.Assert.IsTrue(int.TryParse(fetch.Properties.Items[j].ToString(), out total), "The Total element should be an integer.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2851");

                                    // If the schema validation result is true and Total(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2851,
                                        @"[In Total(ItemOperations)] Element Total in ItemOperations command response (section 2.2.2.9) fetch operation, the parent element is Properties (section 2.2.3.132.1).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2852");

                                    // If the schema validation result is true and Total(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2852,
                                        @"[In Total(ItemOperations)] None [Element Total in ItemOperations command response (section 2.2.2.9) fetch operation has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2854");

                                    // If the schema validation result is true and Total(ItemOperations) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2854,
                                        @"[In Total(ItemOperations)] Element Total in ItemOperations command response (section 2.2.2.9)fetch operation, the number allowed is 0...1 (optional).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2853");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R2853
                                    // If the schema validation result is true, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2853,
                                        @"[In Total(ItemOperations)] Element Total in ItemOperations command response (section 2.2.2.9) fetch operation, the data type is integer ([MS-ASDTYPE] section 2.6).");

                                    this.VerifyIntegerDataType();
                                }
                                #endregion
                            }
                        }
                        #endregion
                    }
                }

                #endregion
            }
            #endregion
        }
        #endregion

        #region Capture code for MeetingResponse command
        /// <summary>
        /// This method is used to verify the MeetingResponse response related requirements.
        /// </summary>
        /// <param name="meetingResponse">MeetingResponse command response.</param>
        private void VerifyMeetingResponseCommand(MeetingResponseResponse meetingResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(meetingResponse.ResponseData, "The MeetingResponse element should not be null.");

            #region Capture code for MeetingResponse
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3312");

            // If the schema validation result is true and MeetingResponse is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3312,
                @"[In MeetingResponse] The MeetingResponse element is a required element in MeetingResponse command requests and MeetingResponse command responses that identifies the body of the HTTP POST as containing a MeetingResponse command (section 2.2.2.9).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1888");

            // If the schema validation result is true and MeetingResponse is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1888,
                @"[In MeetingResponse] None [Element MeetingResponse in MeetingResponse command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1889");

            // If the schema validation result is true and MeetingResponse is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1889,
                @"[In MeetingResponse] Element MeetingResponse in MeetingResponse command response, the child element is Result (section 2.2.3.146.1).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1890");

            // If the schema validation result is true and MeetingResponse is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1890,
                @"[In MeetingResponse] Element MeetingResponse in MeetingResponse command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1891");

            // If the schema validation result is true and MeetingResponse is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1891,
                @"[In MeetingResponse] Element MeetingResponse in MeetingResponse command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Result
            Site.Assert.IsTrue(meetingResponse.ResponseData.Result != null && meetingResponse.ResponseData.Result.Length > 0, "The Result should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3840");

            // If the schema validation result is true and Result(MeetingResponse) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3840,
                @"[In Result(MeetingResponse)] The Result element is a required child element of the MeetingResponse element in MeetingResponse command responses that serves as a container for elements that are sent to the client in the response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2508");

            // If the schema validation result is true and Result(MeetingResponse) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2508,
                @"[In Result(MeetingResponse)] Element Result in MeetingResponse command response (section 2.2.2.10), the parent element is MeetingResponse (section 2.2.3.100).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2509");

            // If the schema validation result is true and Result(MeetingResponse) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2509,
                @"[In Result(MeetingResponse)] Element Result in MeetingResponse command response (section 2.2.2.9), the child elements are RequestId (section 2.2.3.142), Status (section 2.2.3.167.8), CalendarId (section 2.2.3.18), InstanceId (section 2.2.3.87.1).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2510");

            // If the schema validation result is true and Result(MeetingResponse) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2510,
                @"[In Result(MeetingResponse)] Element Result in MeetingResponse command response (section 2.2.2.9), the data type is container ([MS-ASDTYPE] section 2.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2511");

            // If the schema validation result is true and Result(MeetingResponse) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2511,
                @"[In Result(MeetingResponse)] Element Result in MeetingResponse command response (section 2.2.2.9), the number allowed is 1...N (required).");

            foreach (MeetingResponseResult result in meetingResponse.ResponseData.Result)
            {
                #region Capture code for RequestId
                if (result.RequestId != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2467");

                    // If the schema validation result is true and RequestId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2467,
                        @"[In RequestId] Element RequestId in MeetingResponse command response, the parent element is Result (section 2.2.3.142.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2468");

                    // If the schema validation result is true and RequestId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2468,
                        @"[In RequestId] None [Element RequestId in MeetingResponse command response has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2469");

                    // If the schema validation result is true and RequestId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2469,
                        @"[In RequestId] Element RequestId in MeetingResponse command response , the data type is string.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2470");

                    // If the schema validation result is true and RequestId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2470,
                        @"[In RequestId] Element RequestId in MeetingResponse command response , the number allowed is 0…1 (optional).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5871");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5871
                    Site.CaptureRequirementIfIsTrue(
                        result.RequestId.Length <= 64,
                        5871,
                        @"[In RequestId] The RequestId element value is not larger than 64 characters in length.");

                    this.VerifyStringDataType();
                }
                #endregion

                #region Capture code for CalendarID
                if (result.CalendarId != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1107");

                    // If the schema validation result is true and CalendarId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1107,
                        @"[In CalendarId] Element CalendarId in MeetingResponse command response (section 2.2.2.10), the parent element is Result (section 2.2.3.146.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1108");

                    // If the schema validation result is true and CalendarId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1108,
                        @"[In CalendarId] None [Element CalendarId in  MeetingResponse command response (section 2.2.2.10) has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1109");

                    // If the schema validation result is true and CalendarId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1109,
                        @"[In CalendarId] Element CalendarId in MeetingResponse command response (section 2.2.2.10), the data type  is string ([MS-ASDTYPE] section 2.7).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1110");

                    // If the schema validation result is true and CalendarId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1110,
                        @"[In CalendarId] Element CalendarId in MeetingResponse command response (section 2.2.2.10), the number allowed is 0...1 (optional).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5868");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5868
                    Site.CaptureRequirementIfIsTrue(
                        result.CalendarId.Length <= 64,
                        5868,
                        @"[In CalendarId] The CalendarId element value is not larger than 64 characters in length.");

                    this.VerifyStringDataType();
                }
                #endregion

                #region Capture code for Status
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4715");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    4175,
                    @"[In Status(MeetingResponse)] The Status element is a required child element of the Result element in MeetingResponse command responses that indicates the success or failure of the MeetingResponse command request (section 2.2.2.10).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2731");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2731,
                    @"[In Status(MeetingResponse)] Element Status in MeetingResponse command response (section 2.2.2.10),the parent eleemnt is Result (section 2.2.3.146.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2732");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2732,
                    @"[In Status(MeetingResponse)] None [Element Status in MeetingResponse command response (section 2.2.2.10) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2733");

                // Verify MS-ASCMD requirement: MS-ASCMD_R2733
                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2733,
                    @"[In Status(MeetingResponse)] Element Status in MeetingResponse command response (section 2.2.2.10), the data type is integer ([MS-ASDTYPE] section 2.6).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2734");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2734,
                    @"[In Status(MeetingResponse)] Element Status in MeetingResponse command response (section 2.2.2.10), the number allowed is 1…1 (required).");

                Common.VerifyActualValues("Status(MeetingResponse)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4" }), result.Status.ToString(), this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4177");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4177
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4177,
                    @"[In Status(MeetingResponse)] The following table lists the status codes [1,2,3,4] for the MeetingResponse command (section 2.2.2.10). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

                this.VerifyIntegerDataType();
                #endregion
            }
            #endregion
        }
        #endregion

        #region Capture code for MoveItems command
        /// <summary>
        /// This method is used to verify the MoveItems response related requirements.
        /// </summary>
        /// <param name="moveItemsResponse">MoveItems command response.</param>
        private void VerifyMoveItemsCommand(Microsoft.Protocols.TestSuites.Common.MoveItemsResponse moveItemsResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(moveItemsResponse.ResponseData, "The MoveItems element should not be null.");

            #region Capture code for MoveItems
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3476");

            // If the schema validation result is true and MoveItems is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3476,
                @"[In MoveItems] The MoveItems element is a required element in MoveItems command requests and responses that identifies the body of the HTTP POST as containing a MoveItems command (section 2.2.2.11).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1975");

            // If the schema validation result is true and MoveItems is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1975,
                @"[In MoveItems] None [Element MoveItems in MoveItems command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1976");

            // If the schema validation result is true and MoveItems is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1976,
                @"[In MoveItems] Element MoveItems in MoveItems command response, the child element is Response (section 2.2.3.144.4).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1977");

            // If the schema validation result is true and MoveItems is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1977,
                @"[In MoveItems] Element MoveItems in MoveItems command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1978");

            // If the schema validation result is true and MoveItems is not null, this requirement can be verified.
            Site.CaptureRequirement(
                1978,
                @"[In MoveItems] Element MoveItems in MoveItems command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Response
            Site.Assert.IsTrue(moveItemsResponse.ResponseData.Response != null && moveItemsResponse.ResponseData.Response.Length > 0, "The Response should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3822");

            // If the schema validation result is true and Response(MoveItems) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3822,
                @"[In Response(MoveItems)] The Response element is a required child element of the MoveItems element in MoveItems command responses that serves as a container for elements that describe the moved items.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2492");

            // If the schema validation result is true and Response(MoveItems) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2492,
                @"[In Response(MoveItems)] Element Response in MoveItems command response (section 2.2.2.11),the parent element is MoveItems (section 2.2.3.111.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2493");

            // If the schema validation result is true and Response(MoveItems) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2493,
                @"[In Response(MoveItems)] Element Response in MoveItems command response (section 2.2.2.11), the child element are SrcMsgId (section 2.2.3.165), Status (section 2.2.3.167.9), DstMsgId (section 2.2.3.50).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2494");

            // If the schema validation result is true and Response(MoveItems) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2494,
                @"[In Response(MoveItems)] Element Response in MoveItems command response (section 2.2.2.11), the data type is container ([MS-ASDTYPE] section 2.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2495");

            // If the schema validation result is true and Response(MoveItems) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2495,
                @"[In Response(MoveItems)] Element Response in MoveItems command response (section 2.2.2.11), the number allowed is 1…N (required).");

            foreach (TestSuites.Common.Response.MoveItemsResponse response in moveItemsResponse.ResponseData.Response)
            {
                #region Capture code for SrcMsgId
                Site.Assert.IsNotNull(response.SrcMsgId, "The SrcMsgId should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2675");

                // If the schema validation result is true and SrcMsgId is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2675,
                    @"[In SrcMsgId] Element SrcMsgId in MoveItems command response, the parent element is Response (section 2.2.3.144.4).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2676");

                // If the schema validation result is true and SrcMsgId is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2676,
                    @"[In SrcMsgId] None [Element SrcMsgId in MoveItems command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2677");

                // If the schema validation result is true and SrcMsgId is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2677,
                    @"[In SrcMsgId] Element SrcMsgId in MoveItems command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2678");

                // If the schema validation result is true and SrcMsgId is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2678,
                    @"[In SrcMsgId] Element SrcMsgId in MoveItems command response, the number allowed is 1…1 (required).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5882");

                // Verify MS-ASCMD requirement: MS-ASCMD_R5882
                Site.CaptureRequirementIfIsTrue(
                    response.SrcMsgId.Length <= 64,
                    5882,
                    @"[In SrcMsgId] The SrcMsgId element value is not larger than 64 characters in length.");

                this.VerifyStringDataType();
                #endregion

                #region Capture code for DstMsgId
                if (response.DstMsgId != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1571");

                    // If the schema validation result is true and DstMsgId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1571,
                        @"[In DstMsgId] Element DstMsgId in MoveItems command response (section 2.2.2.10), the parent element is Response (section 2.2.3.140.4).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1572");

                    // If the schema validation result is true and DstMsgId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1572,
                        @"[In DstMsgId] None [Element DstMsgId in MoveItems command response (section 2.2.2.10) has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1573");

                    // If the schema validation result is true and DstMsgId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1573,
                        @"[In DstMsgId] Element DstMsgId in MoveItems command response (section 2.2.2.10), the data type is string ([MS-ASDTYPE] section 2.6).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1574");

                    // If the schema validation result is true and DstMsgId is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1574,
                        @"[In DstMsgId] Element DstMsgId in MoveItems command response (section 2.2.2.10), the number allowed is 0…1 (optional).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5870");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R5870
                    Site.CaptureRequirementIfIsTrue(
                        response.DstMsgId.Length <= 64,
                        5870,
                        @"[In DstMsgId] The DstMsgId element value is not larger than 64 characters in length.");

                    this.VerifyStringDataType();
                }
                #endregion

                #region Capture code for Status
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4200");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    4200,
                    @"[In Status(MoveItems)] The Status element is a required child element of the Response element in MoveItems command responses that indicates the success or failure the MoveItems command request (section 2.2.2.11).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2735");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2735,
                    @"[In Status(MoveItems)] Element Status in MoveItems command response, the parent element is Response (section 2.2.3.144.4).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2738");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2738,
                    @"[In Status(MoveItems)] Element Status in MoveItems command response, the number allowed is 1…1 (required).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2737");

                // Verify MS-ASCMD requirement: MS-ASCMD_R2737
                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2737,
                    @"[In Status(MoveItems)] Element Status in MoveItems command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

                Common.VerifyActualValues("Status(MoveItems)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "7" }), response.Status.ToString(), this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4203");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4203
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4203,
                    @"[In Status(MoveItems)] The following table lists the status codes [1,2,3,4,5,7] for the MoveItems command (section 2.2.2.11). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

                this.VerifyIntegerDataType();
                #endregion
            }
            #endregion
        }
        #endregion

        #region Capture code for Ping command
        /// <summary>
        /// This method is used to verify the Ping response related requirements.
        /// </summary>
        /// <param name="pingResponse">Ping command response.</param>
        private void VerifyPingCommand(PingResponse pingResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(pingResponse.ResponseData, "The Ping element should not be null.");

            #region Capture code for Ping
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3676");

            // If the schema validation result is true and Ping is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3676,
                @"[In Ping] The Ping element is a required element in Ping command requests and responses that identifies the body of the HTTP POST as containing a Ping command (section 2.2.2.12).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2388");

            // If the schema validation result is true and Ping is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2388,
                @"[In Ping] None [Element Ping in Ping command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2389");

            // If the schema validation result is true and Ping is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2389,
                @"[In Ping] Element Ping in Ping command response, the child elements are HeartbeatInterval, Folders, MaxFolders (section 2.2.3.96), Status (section 2.2.3.167.10).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2390");

            // If the schema validation result is true and Ping is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2390,
                @"[In Ping] Element Ping in Ping command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2391");

            // If the schema validation result is true and Ping is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2391,
                @"[In Ping] Element Ping in Ping command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Status
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4227");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                4227,
                @"[In Status(Ping)] The Status element is a required child element of the Ping element in Ping command responses that indicates the success or failure of the Ping command request (section 2.2.2.11).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2739");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2739,
                @"[In Status(Ping)] Element Status in Ping command response, the parent element is Ping (section 2.2.3.130).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2740");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2740,
                @"[In Status(Ping)] None [Element Status in Ping command response has no child element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2742");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2742,
                @"[In Status(Ping)] Element Status in Ping command response, the number allowed is 1…1 (required).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2741");

            // Verify MS-ASCMD requirement: MS-ASCMD_R2741
            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                2741,
                @"[In Status(Ping)] Element Status in Ping command response, the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            Common.VerifyActualValues("Status(Ping)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "7", "8" }), pingResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4231");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4231
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4231,
                @"[In Status(Ping)] The following table lists the status codes [1,2,3,4,5,6,7,8] for the Ping command (section 2.2.2.12). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

            this.VerifyIntegerDataType();
            #endregion

            #region Capture code for HeartbeatInterval
            if (pingResponse.ResponseData.HeartbeatInterval != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1780");

                // If the schema validation result is true and HeartbeatInterval(Ping) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1780,
                    @"[In HeartbeatInterval(Ping)] Element HeartbeatInterval in Ping command response, the parent element is Ping.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1781");

                // If the schema validation result is true and HeartbeatInterval(Ping) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1781,
                    @"[In HeartbeatInterval(Ping)] None [ement HeartbeatInterval in Ping command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1783");

                // If the schema validation result is true and HeartbeatInterval(Ping) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1783,
                    @"[In HeartbeatInterval(Ping)] Ement HeartbeatInterval in Ping command response, the number allowed is 0...1 (optional).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1782");

                // Verify MS-ASCMD requirement: MS-ASCMD_R1782
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(int),
                    Convert.ToInt32(pingResponse.ResponseData.HeartbeatInterval).GetType(),
                    1782,
                    @"[In HeartbeatInterval(Ping)] Element HeartbeatInterval in Ping command response, the data type is integer.");

                this.VerifyIntegerDataType();
            }
            #endregion

            #region Capture code for Folders
            if (pingResponse.ResponseData.Folders != null && pingResponse.ResponseData.Folders.Length > 0)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1707");

                // If the schema validation result is true and Folders is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1707,
                    @"[In Folders(Ping)] Element Folders in Ping command response, the parent element is Ping.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1708");

                // If the schema validation result is true and Folders is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1708,
                    @"[In Folders(Ping)] Element Folders in Ping command response, the child element is Folder.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1709");

                // If the schema validation result is true and Folders is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1709,
                    @"[In Folders(Ping)] Element Folders in Ping command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1710");

                // If the schema validation result is true and Folders is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1710,
                    @"[In Folders(Ping)] Element Folders in Ping command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                foreach (string folder in pingResponse.ResponseData.Folders)
                {
                    #region Capture code for Folder
                    Site.Assert.IsNotNull(folder, "The Folder element of the Folders element should not be null.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3082");

                    // If the schema validation result is true and Folder is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        3082,
                        @"[In Folders(Ping)] The Folder element is a required child element of the Folders element in Ping command responses that identifies the folder that is being described by the returned status code.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1675");

                    // If the schema validation result is true and Folder is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1675,
                        @"[In Folders(Ping)] Element Folder in Ping command response (section 2.2.2.12), the parent element is Folders (section 2.2.3.70).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1676");

                    // If the schema validation result is true and Folder is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1676,
                        @"[In Folders(Ping)] None [Element Folder in Ping command response has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1677");

                    // If the schema validation result is true and Folder is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1677,
                        @"[In Folders(Ping)] Element Folder in Ping command response, the data type is string ([MS-ASDTYPE] section 2.7).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1678");

                    // If the schema validation result is true and Folder is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1678,
                        @"[In Folders(Ping)] Element Folder in Ping command response, the number allowed is 1...N (required).");

                    this.VerifyStringDataType();
                    #endregion
                }
            }
            #endregion
        }
        #endregion

        #region Capture code for Provision command
        /// <summary>
        /// This method is used to verify the Provision response related requirements.
        /// </summary>
        /// <param name="provisionResponse">Provision command response.</param>
        private void VerifyProvisionCommand(ProvisionResponse provisionResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(provisionResponse.ResponseData, "The Provision element should not be null.");

            if (provisionResponse.ResponseData.Policies != null && provisionResponse.ResponseData.Policies.Policy != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4996");

                // If the schema validation result is true and Policy is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    4996,
                    @"[In Downloading Policy Settings] The response from the server contains provision:PolicyType, provision:PolicyKey, and provision:Status elements.");
            }
        }
        #endregion

        #region Capture code for ResolveRecipients command
        /// <summary>
        /// This method is used to verify the ResolveRecipients response related requirements.
        /// </summary>
        /// <param name="resolveRecipientsResponse">ResolveRecipients command response.</param>
        private void VerifyResolveRecipientsCommand(Microsoft.Protocols.TestSuites.Common.ResolveRecipientsResponse resolveRecipientsResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(resolveRecipientsResponse.ResponseData, "The ResolveRecipients element should not be null.");

            #region Capture code for ResolveRecipients
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3809");

            // If the schema validation result is true and ResolveRecipients is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3809,
                @"[In ResolveRecipients] The ResolveRecipients element is a required element in ResolveRecipients command requests and responses that identifies the body of the HTTP POST as containing a ResolveRecipients command (section 2.2.2.13).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2475");

            // If the schema validation result is true and ResolveRecipients is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2475,
                @"[In ResolveRecipients] None [Element ResolveRecipients in ResolveRecipients command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2476");

            // If the schema validation result is true and ResolveRecipients is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2476,
                @"[In ResolveRecipients] Element ResolveRecipients in ResolveRecipients command response, the child elements are Status (section 2.2.3.167.11), Response (section 2.2.3.144.5).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2477");

            // If the schema validation result is true and ResolveRecipients is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2477,
                @"[In ResolveRecipients] Element ResolveRecipients in ResolveRecipients command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2478");

            // If the schema validation result is true and ResolveRecipients is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2478,
                @"[In ResolveRecipients] Element ResolveRecipients in ResolveRecipients command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            #endregion

            #region Capture code for Status
            this.VerifyStatusElementForResolveRecipients(int.Parse(resolveRecipientsResponse.ResponseData.Status));

            Common.VerifyActualValues("Status(ResolveRecipients)", AdapterHelper.ValidStatus(new string[] { "1", "5", "6" }), resolveRecipientsResponse.ResponseData.Status.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4267");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4267
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4267,
                @"[In Status(ResolveRecipients)] The following table shows valid values [1,5,6] for the Status element when it is returned as a child of the ResolveRecipients element.");
            #endregion

            #region Capture code for Response
            if (resolveRecipientsResponse.ResponseData.Response != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2496");

                // If the schema validation result is true and Response(ResolveRecipients) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2496,
                    @"[In Response(ResolveRecipients)] Element Response in ResolveRecipients command response (section 2.2.2.14), the parent element is ResolveRecipients (section 2.2.3.143)");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2497");

                // If the schema validation result is true and Response(ResolveRecipients) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2497,
                    @"[In Response(ResolveRecipients)] Element Response in ResolveRecipients command response (section 2.2.2.14), the child element is To (section 2.2.3.173), Status (section 2.2.3.167.11), RecipientCount (section 2.2.3.137), Recipient (section 2.2.3.136).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2498");

                // If the schema validation result is true and Response(ResolveRecipients) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2498,
                    @"[In Response(ResolveRecipients)] Element Response in ResolveRecipients command response (section 2.2.2.14), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2499");

                // If the schema validation result is true and Response(ResolveRecipients) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2499,
                    @"[In Response(ResolveRecipients)] Element Response in ResolveRecipients command response (section 2.2.2.14), the number allowed is 0…1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Status
                int status;

                Site.Assert.IsNotNull(resolveRecipientsResponse.ResponseData.Response[0].Status, "As a child element of the Response element, the Status element should not be null.");
                Site.Assert.IsTrue(int.TryParse(resolveRecipientsResponse.ResponseData.Response[0].Status, out status), "The Status element should be a byte.");

                this.VerifyStatusElementForResolveRecipients(status);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4273");

                // If the schema validation result is true and Status(ResolveRecipients) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    4273,
                    @"[In Status(ResolveRecipients)] As a child element of the Response element, the Status element provides the status of the ResolveRecipients command response Response element.");

                Common.VerifyActualValues("Status(ResolveRecipients)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4" }), resolveRecipientsResponse.ResponseData.Response[0].Status, this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4274");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4274
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4274,
                    @"[In Status(ResolveRecipients)] The following table shows valid values [1,2,3,4] for the Status element when it is returned as a child element of the Response element.");
                #endregion

                #region Capture code for To
                Site.Assert.IsNotNull(resolveRecipientsResponse.ResponseData.Response[0].To, "The To element should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4609");

                // If the schema validation result is true and To is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    4609,
                    @"[In To] The To element is a required child element of the Response element in ResolveRecipients command responses that specifies a recipient to be resolved.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2847");

                // If the schema validation result is true and To is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2847,
                    @"[In To] Element To in ResolveRecipients command response, the parent element is Response (section 2.2.3.144.5).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2848");

                // If the schema validation result is true and To is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2848,
                    @"[In To] None [Element To in ResolveRecipients command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2849");

                // If the schema validation result is true and To is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2849,
                    @"[In To] Element To in ResolveRecipients command response, the data type is string.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2850");

                // If the schema validation result is true and To is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2850,
                    @"[In To] Element To in ResolveRecipients command response, the number allowed is 1…1 (required).");
                #endregion

                #region Capture code for RecipientCount
                if (resolveRecipientsResponse.ResponseData.Response[0].RecipientCount != null)
                {
                    this.VerifyRecipientCountElement(Convert.ToInt32(resolveRecipientsResponse.ResponseData.Response[0].RecipientCount));

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3767");

                    // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        3767,
                        @"[In RecipientCount] As a child element of the Response element, the RecipientCount element specifies the number of recipients that are returned in the ResolveRecipients command response.");
                }
                #endregion

                #region Capture code for Recipient
                if (resolveRecipientsResponse.ResponseData.Response[0].Recipient != null && resolveRecipientsResponse.ResponseData.Response[0].Recipient.Length > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2431");

                    // If the schema validation result is true and Recipient is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2431,
                        @"[In Recipient] Element Recipient in ResolveRecipients command response (section 2.2.2.14), the parent element is Response (section 2.2.3.144.5).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2432");

                    // If the schema validation result is true and Recipient is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2432,
                        @"[In Recipient] Element Recipient in ResolveRecipients command response (section 2.2.2.14), the child elements are Type (section 2.2.3.176.4), DisplayName (section 2.2.3.47.5), EmailAddress (section 2.2.3.53.2), Availability (section 2.2.3.16), Certificates (section 2.2.3.23.1), Picture (section 2.2.3.129.1).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2433");

                    // If the schema validation result is true and Recipient is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2433,
                        @"[In Recipient] Element Recipient in ResolveRecipients command response (section 2.2.2.14), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2434");

                    // If the schema validation result is true and Recipient is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2434,
                        @"[In Recipient] Element Recipient in ResolveRecipients command response (section 2.2.2.14), the number allowed is 0...N (optional).");

                    this.VerifyContainerDataType();

                    foreach (ResolveRecipientsResponseRecipient recipient in resolveRecipientsResponse.ResponseData.Response[0].Recipient)
                    {
                        #region Capture code for EmailAddress
                        Site.Assert.IsNotNull(recipient.EmailAddress, "The EmailAddress element should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5309");

                        // If the schema validation result is true and EmailAddress(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5309,
                            @"[In EmailAddress(ResolveRecipients)] The EmailAddress element is a required child element of the Recipient element in ResolveRecipients command responses that contains the email address of the recipient, in SMTP format.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5311");

                        // If the schema validation result is true and EmailAddress(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5311,
                            @"[In EmailAddress(ResolveRecipients)] Element EmailAddress in ResolveRecipients command response (section 2.2.2.13), the parent element is Recipient (section 2.2.3.132).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5312");

                        // If the schema validation result is true and EmailAddress(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5312,
                            @"[In EmailAddress(ResolveRecipients)] None [Element EmailAddress in ResolveRecipients command response (section 2.2.2.13)has no child element]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5313");

                        // If the schema validation result is true and EmailAddress(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5313,
                            @"[In EmailAddress(ResolveRecipients)] Element EmailAddress in ResolveRecipients command response (section 2.2.2.13), the data type is string ([MS-ASDTYPE] section 2.6).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5314");

                        // If the schema validation result is true and EmailAddress(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            5314,
                            @"[In EmailAddress(ResolveRecipients)] Element EmailAddress in ResolveRecipients command response (section 2.2.2.13), the number allowed is 1…1 (required).");

                        this.VerifyStringDataType();
                        #endregion

                        #region Capture code for Type
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4686");

                        // If the schema validation result is true and Type(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            4686,
                            @"[In Type(ResolveRecipients)] The Type element is a required child element of the Recipient element in ResolveRecipients command responses that indicates the type of recipient, either a contact entry (2) or a GAL entry (1).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2875");

                        // If the schema validation result is true and Type(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2875,
                            @"[In Type(ResolveRecipients)] Element Type in ResolveRecipients command response (section 2.2.2.14),the parent element is Recipient (section 2.2.3.136).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2876");

                        // If the schema validation result is true and Type(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2876,
                            @"[In Type(ResolveRecipients)] None [Element Type in ResolveRecipients command response (section 2.2.2.14)has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2877");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2877
                        Site.CaptureRequirementIfAreEqual<Type>(
                            typeof(byte),
                            recipient.Type.GetType(),
                            2877,
                            @"[In Type(ResolveRecipients)] Element Type in ResolveRecipients command response (section 2.2.2.14), the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

                        this.VerifyUnsignedByteDataType(recipient.Type);

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2878");

                        // If the schema validation result is true and Type(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2878,
                            @"[In Type(ResolveRecipients)] Element Type in ResolveRecipients command response (section 2.2.2.14), the number allowed is 1...1 (required).");
                        #endregion

                        #region Capture code for DisplayName
                        Site.Assert.IsNotNull(recipient.DisplayName, "The DisplayName element should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2199");

                        // If the schema validation result is true and DisplayName(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2199,
                            @"[In DisplayName(ResolveRecipients)] The DisplayName element is a required child element of the Recipient element in ResolveRecipients command responses that contains the display name of the recipient.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1543");

                        // If the schema validation result is true and DisplayName(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1543,
                            @"[In DisplayName(ResolveRecipients)] Element DisplayName in ResolveRecipients command response (section 2.2.2.14), the parent element is Recipient (section 2.2.3.136).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1544");

                        // If the schema validation result is true and DisplayName(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1544,
                            @"[In DisplayName(ResolveRecipients)] None [Element DisplayName in ResolveRecipients command response (section 2.2.2.14) has no child element .]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1545");

                        // If the schema validation result is true and DisplayName(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1545,
                            @"[In DisplayName(ResolveRecipients)] Element DisplayName in ResolveRecipients command response (section 2.2.2.14), the data type is string ([MS-ASDTYPE] section 2.7).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1546");

                        // If the schema validation result is true and DisplayName(ResolveRecipients) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1546,
                            @"[In DisplayName(ResolveRecipients)] Element DisplayName in ResolveRecipients command response (section 2.2.2.14), the number allowed is 1...1 (required).");

                        this.VerifyStringDataType();
                        #endregion

                        #region Capture code for Availablity
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && recipient.Availability != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1091");

                            // If the schema validation result is true and Availability is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1091,
                                @"[In Availability] Element Availability in ResolveRecipients command response, the parent element is Recipient (section 2.2.3.136).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1092");

                            // If the schema validation result is true and Availability is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1092,
                                @"[In Availability] Element Availability in ResolveRecipients command response, the child elements are Status (section 2.2.3.167.11), MergedFreeBusy (section 2.2.3.101).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1093");

                            // If the schema validation result is true and Availability is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1093,
                                @"[In Availability] Element Availability in ResolveRecipients command response, the data type is container.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1094");

                            // If the schema validation result is true and Availability is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1094,
                                @"[In Availability] Element Availability in ResolveRecipients command response, the number allowed is 0...1 (optional).");

                            this.VerifyContainerDataType();

                            #region Capture code for Status
                            Site.Assert.IsNotNull(recipient.Availability.Status, "As a child element of Availability, the Status should not be null.");

                            byte availabilityStatus;

                            Site.Assert.IsTrue(byte.TryParse(recipient.Availability.Status, out availabilityStatus), "As a child element of Availability, the Status should be a byte.");

                            this.VerifyStatusElementForResolveRecipients(availabilityStatus);

                            Common.VerifyActualValues("Status(ResolveRecipients)", AdapterHelper.ValidStatus(new string[] { "1", "160", "161", "162", "163" }), recipient.Availability.Status, this.Site);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4290");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R4290
                            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                            Site.CaptureRequirement(
                                4290,
                                @"[In Status(ResolveRecipients)] The following table shows valid values [1,160,161,162,163] for the Status element when it is returned as a child element of the Availability element.");
                            #endregion

                            #region Capture code for MergedFreeBusy
                            if (recipient.Availability.MergedFreeBusy != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3315");

                                // If the schema validation result is true and MergedFreeBusy is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    3315,
                                    @"[In MergedFreeBusy] The MergedFreeBusy element<49> is an optional child element of the Availability element in ResolveRecipients command responses that specifies the free/busy information for the users or distribution list identified in the request.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1892");

                                // If the schema validation result is true and MergedFreeBusy is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1892,
                                    @"[In MergedFreeBusy] Element MergedFreeBusy in ResolveRecipients command response (section 2.2.2.14), the parent element is Availability (section 2.2.3.16).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1893");

                                // If the schema validation result is true and MergedFreeBusy is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1893,
                                    @"[In MergedFreeBusy] None [Element MergedFreeBusy in ResolveRecipients command response (section 2.2.2.14) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1894");

                                // If the schema validation result is true and MergedFreeBusy is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1894,
                                    @"[In MergedFreeBusy] Element MergedFreeBusy in ResolveRecipients command response (section 2.2.2.14),  the data type is string ([MS-ASDTYPE] section 2.7).");

                                this.VerifyStringDataType();

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1895");

                                // If the schema validation result is true and MergedFreeBusy is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1895,
                                    @"[In MergedFreeBusy] Element MergedFreeBusy in ResolveRecipients command response (section 2.2.2.14), the number allowed is 0...1 (optional).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3317");

                                // Verify MS-ASCMD requirement: MS-ASCMD_R3317
                                Site.CaptureRequirementIfIsTrue(
                                    recipient.Availability.MergedFreeBusy.Length <= 32768,
                                    3317,
                                    @"[In MergedFreeBusy] The MergedFreeBusy element value string has a maximum length of 32 KB.");

                                Regex mergedFreeBusyRegex = new Regex("^[0-4]{0,32768}$");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3320");

                                // Verify MS-ASCMD requirement: MS-ASCMD_R3320
                                Site.CaptureRequirementIfIsTrue(
                                    mergedFreeBusyRegex.IsMatch(recipient.Availability.MergedFreeBusy),
                                    3320,
                                    @"[In MergedFreeBusy] The following table lists the valid values[0, 1, 2, 3, 4]].");
                            }
                            #endregion
                        }
                        #endregion

                        #region Capture code for Picture
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0") && recipient.Picture != null && recipient.Picture.Length > 0)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2372");

                            // If the schema validation result is true and Picture(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2372,
                                @"[In Picture(ResolveRecipients)] Element Picture in ResolveRecipients command response, the parent element is Recipient (section 2.2.3.136).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2373");

                            // If the schema validation result is true and Picture(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2373,
                                @"[In Picture(ResolveRecipients)] Element Picture in ResolveRecipients command response, the child elements are Status (section 2.2.3.167.11), Data (section 2.2.3.39.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2374");

                            // If the schema validation result is true and Picture(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2374,
                                @"[In Picture(ResolveRecipients)] Element Picture in ResolveRecipients command response, the data type is container.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2375");

                            // If the schema validation result is true and Picture(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2375,
                                @"[In Picture(ResolveRecipients)] Element Picture in ResolveRecipients command response, the number allowed is 0…1 (optional).");

                            this.VerifyContainerDataType();

                            foreach (ResolveRecipientsResponseRecipientPicture picture in recipient.Picture)
                            {
                                #region Capture code for Status
                                Site.Assert.IsNotNull(picture.Status != null, "As a child element of Picture, the Status should not be null.");

                                byte pictureStatus;

                                Site.Assert.IsTrue(byte.TryParse(picture.Status, out pictureStatus), "As a child element of Picture, the Status should be a byte.");

                                this.VerifyStatusElementForResolveRecipients(pictureStatus);

                                Common.VerifyActualValues("Status(ResolveRecipients)", AdapterHelper.ValidStatus(new string[] { "1", "173", "174", "175" }), picture.Status, this.Site);

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4309");

                                // Verify MS-ASCMD requirement: MS-ASCMD_R4309
                                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                                Site.CaptureRequirement(
                                    4309,
                                    @"[In Status(ResolveRecipients)] The following table shows valid values [1,173,174,175] for the Status element when it is returned as a child element of the Picture element.");
                                #endregion

                                #region Capture code for Data
                                if (picture.Data != null)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1466");

                                    // If the schema validation result is true and Data(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1466,
                                        @"[In Data(ResolveRecipients)] Element Data in ResolveRecipients command response (section 2.2.2.14), the parent element is Picture (section 2.2.3.129.1).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1467");

                                    // If the schema validation result is true and Data(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1467,
                                        @"[In Data(ResolveRecipients)] None [Element Data in ResolveRecipients command response (section 2.2.2.14) has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1468");

                                    // If the schema validation result is true and Data(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1468,
                                        @"[In Data(ResolveRecipients)] Element Data in ResolveRecipients command response  (section 2.2.2.14), the data type is Byte array.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1469");

                                    // If the schema validation result is true and Data(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1469,
                                        @"[In Data(ResolveRecipients)] Element Data in ResolveRecipients command response  (section 2.2.2.14), the number allowed is 0…1 (optional).");
                                }
                                #endregion
                            }
                        }
                        #endregion

                        #region Capture code for Certificates
                        if (recipient.Certificates != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1135");

                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1135,
                                @"[In Certificates(ResolveRecipients)] Element Certificates(ResolveRecipients) in ResolveRecipients command response (section 2.2.2.14), the parent element is Recipient (section 2.2.3.136).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1136");

                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1136,
                                @"[In Certificates(ResolveRecipients)] Element Certificates(ResolveRecipients) in ResolveRecipients command response (section 2.2.2.14), the child elements are Status (section2.2.3.167.11), CertificateCount (section 2.2.3.21) , RecipientCount (section 2.2.3.137) , Certificate (section 2.2.3.19) , MiniCertificate (section 2.2.3.106)");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1137");

                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1137,
                                @"[In Certificates(ResolveRecipients)] Element Certificates(ResolveRecipients) in ResolveRecipients command response (section 2.2.2.14), the data type is container ([MS-ASDTYPE] section 2.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1138");

                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1138,
                                @"[In Certificates(ResolveRecipients)] Element Certificates(ResolveRecipients) in ResolveRecipients command response (section 2.2.2.14), the number allowed is  0...1 (optional).");

                            this.VerifyContainerDataType();

                            #region Capture code for Status
                            int certificatesStatus = int.Parse(recipient.Certificates.Status);

                            this.VerifyStatusElementForResolveRecipients(certificatesStatus);

                            Common.VerifyActualValues("Status(ResolveRecipients)", AdapterHelper.ValidStatus(new string[] { "1", "7", "8" }), certificatesStatus.ToString(), this.Site);

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4301");

                            // Verify MS-ASCMD requirement: MS-ASCMD_R4301
                            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                            Site.CaptureRequirement(
                                4301,
                                @"[In Status(ResolveRecipients)] The following table shows valid values [1,7,8] for the Status element when it is returned as a child element of the Certificates element.");
                            #endregion

                            #region Capture code for CertificateCount
                            if (recipient.Certificates.CertificateCount != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1127");

                                // If the schema validation result is true and CertificateCount is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1127,
                                    @"[In CertificateCount] Element CertificateCount in ResolveRecipients command response (section 2.2.2.14), the parent element is Certificates (section 2.2.3.23.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1128");

                                // If the schema validation result is true and CertificateCount is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1128,
                                    @"[In CertificateCount] None [Element CertificateCount in ResolveRecipients command response (section 2.2.2.14) has no child element. ]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1129");

                                // Verify MS-ASCMD requirement: MS-ASCMD_R1129
                                Site.CaptureRequirementIfAreEqual<Type>(
                                    typeof(int),
                                    Convert.ToInt32(recipient.Certificates.CertificateCount).GetType(),
                                    1129,
                                    @"[In CertificateCount] Element CertificateCount in ResolveRecipients command response (section 2.2.2.14), the data type is integer ([MS-ASDTYPE] section 2.6).");

                                this.VerifyIntegerDataType();

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1130");

                                // If the schema validation result is true and CertificateCount is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1130,
                                    @"[In CertificateCount] Element CertificateCount in ResolveRecipients command response (section 2.2.2.14), the number allowed is 0…1 (optional) per Certificates parent element.");
                            }
                            #endregion

                            #region Capture code for RecipientCount
                            if (recipient.Certificates.RecipientCount != null)
                            {
                                this.VerifyRecipientCountElement(Convert.ToInt32(recipient.Certificates.RecipientCount));

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3768");

                                // If the schema validation result is true and RecipientCount is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    3768,
                                    @"[In RecipientCount] As a child element of the Certificates element, the RecipientCount element specifies the number of members belonging to a distribution list.");
                            }
                            #endregion

                            #region Capture code for Certificate
                            if (recipient.Certificates.Certificate != null && recipient.Certificates.Certificate.Length > 0)
                            {
                                for (int i = 0; i < recipient.Certificates.Certificate.Length; i++)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1111");

                                    // If the schema validation result is true and Certificate(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1111,
                                        @"[In Certificate(ResolveRecipients)] Element Certificate in ResolveRecipients command response (section 2.2.2.14), the parent element is Certificates (section 2.2.3.23.1).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1112");

                                    // If the schema validation result is true and Certificate(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1112,
                                        @"[In Certificate(ResolveRecipients)] None [Element Certificate in ResolveRecipients command response (section 2.2.2.14) has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1113");

                                    // If the schema validation result is true and Certificate(ResolveRecipients) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1113,
                                        @"[In Certificate(ResolveRecipients)] Element Certificate in ResolveRecipients command response (section 2.2.2.14), the data type is string ([MS-ASDTYPE] section 2.7) (encoded with base64 encoding).");

                                    this.VerifyStringDataType();
                                }

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1114");

                                // If the schema validation result is true and Certificate(ResolveRecipients) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1114,
                                    @"[In Certificate(ResolveRecipients)] Element Certificate in ResolveRecipients command response (section 2.2.2.14), the number allowed is 0...N (optional).");
                            }
                            #endregion

                            #region Capture code for MiniCertificate
                            if (recipient.Certificates.MiniCertificate != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1939");

                                // If the schema validation result is true and MiniCertificate is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1939,
                                    @"[In MiniCertificate] Element MiniCertificate in ResolveRecipients command response (section 2.2.2.13), the parent element is Certificates (section 2.2.3.23.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1940");

                                // If the schema validation result is true and MiniCertificate is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1940,
                                    @"[In MiniCertificate] None [Element MiniCertificate in ResolveRecipients command response (section 2.2.2.13)has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1941");

                                // If the schema validation result is true and MiniCertificate is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1941,
                                    @"[In MiniCertificate] Element MiniCertificate in ResolveRecipients command response (section 2.2.2.14),  the data type is string ([MS-ASDTYPE] section 2.7) (encoded with base64 encoding).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1942");

                                // If the schema validation result is true and MiniCertificate is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1942,
                                    @"[In MiniCertificate] Element MiniCertificate in ResolveRecipients command response (section 2.2.2.14),  the number allowed is 0...1 per Certificates parent element.");

                                this.VerifyStringDataType();
                            }
                            #endregion
                        }
                        #endregion
                    }
                }
                #endregion
            }
            #endregion
        }
        #endregion

        #region Capture code for Search command
        /// <summary>
        /// This method is used to verify the Search response related requirements.
        /// </summary>
        /// <param name="searchResponse">Search command response.</param>
        private void VerifySearchCommand(Microsoft.Protocols.TestSuites.Common.SearchResponse searchResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(searchResponse.ResponseData, "The Search element should not be null.");

            #region Capture code for Search
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3880");

            // If the schema validation result is true and Search is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3880,
                @"[In Search] The Search element is a required element in Search command requests and responses that identifies the body of the HTTP POST as containing a Search command (section 2.2.2.15).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2545");

            // If the schema validation result is true and Search is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2545,
                @"[In Search] None [Element Search in Search command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2546");

            // If the schema validation result is true and Search is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2546,
                @"[In Search] Element Search in Search command response, the child elements are Status (section 2.2.3.167.12), Response (section 2.2.3.144.6).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2547");

            // If the schema validation result is true and Search is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2547,
                @"[In Search] Element Search in Search command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2548");

            // If the schema validation result is true and Search is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2548,
                @"[In Search] Element Search in Search command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for Status
            Site.Assert.IsNotNull(searchResponse.ResponseData.Status, "The Status element should not be null.");

            int status;

            Site.Assert.IsTrue(int.TryParse(searchResponse.ResponseData.Status, out status), "The Status element should be an integer.");

            this.VerifyStatusElementForSearch();

            Common.VerifyActualValues("Status(Search)", AdapterHelper.ValidStatus(new string[] { "1", "3" }), searchResponse.ResponseData.Status, this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4321");

            // Verify MS-ASCMD requirement: MS-ASCMD_R4321
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                4321,
                @"[In Status(Search)] The following table specifies valid values [1,3] for the Status element when it is returned as a child element of the Search element.");
            #endregion

            #region Capture code for Response
            if (searchResponse.ResponseData.Response != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2500");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2500,
                    @"[In Response(Search)] Element Response in Search command response (section 2.2.2.15), the parent element is Search (section 2.2.3.150).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2501");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2501,
                    @"[In Response(Search)] Element Response in Search command response (section 2.2.2.15), the child element is Store (section 2.2.3.168.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2502");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2502,
                    @"[In Response(Search)] Element Response in Search command response (section 2.2.2.15), the data type is container ([MS-ASDTYPE] section .");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2503");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2503,
                    @"[In Response(Search)] Element Response in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Store
                Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store, "The Store element in Search command response should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4531");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    4531,
                    @"[In Store(Search)] The Store element is a required child element of the Response element in Search command responses that contains the Status, Result, Range, and Total elements for the returned mailbox entries.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2783");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2783,
                    @"[In Store(Search)] Element Store in Search command response, the parent element is Response (section 2.2.3.144.6).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2784");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2784,
                    @"[In Store(Search)] Element Store in Search command response, the child elements are Status (section 2.2.3.167.12), Result (section 2.2.3.146.2), Range (section 2.2.3.134.2), Total (section 2.2.3.174.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2785");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2785,
                    @"[In Store(Search)] Element Store in Search command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2786");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2786,
                    @"[In Store(Search)] Element Store in Search command response, the number allowed is 1...1 (required).");

                #region Capture code for Total
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4633");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    4633,
                    @"[In Total(Search)] The Total element is a required child element of the Store element in Search command responses that provides an estimate of the total number of mailbox entries that matched the search Query element (section 2.2.3.133) value.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2855");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2855,
                    @"[In Total(Search)] Element Total in Search command response (section 2.2.2.15), the parent element is Store (section 2.2.3.168.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2856");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    2856,
                    @"[In Total(Search)] None [Element Total in Search command response (section 2.2.2.15) has no child element.]");

                int total;
                if (searchResponse.ResponseData.Response.Store.Total != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2857");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R2857
                    Site.CaptureRequirementIfIsTrue(
                        int.TryParse(searchResponse.ResponseData.Response.Store.Total, out total),
                        2857,
                        @"[In Total(Search)] Element Total in Search command response (section 2.2.2.15), the data type is integer ([MS-ASDTYPE] section 2.6).");
                }

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2858");

                // If the schema validation result is true and Total(Search) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2858,
                    @"[In Total(Search)] Element Total in Search command response (section 2.2.2.15), the number allowed is 1...1 (required).");
                #endregion

                #region Capture code for Status
                Site.Assert.IsNotNull(searchResponse.ResponseData.Response.Store.Status, "The Status element should not be null.");

                this.VerifyStatusElementForSearch();

                Common.VerifyActualValues("Status", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "10", "11", "12", "13", "14" }), searchResponse.ResponseData.Response.Store.Status, this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4325");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4325
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4325,
                    @"[In Status(Search)] The following table specifies valid values [1,2,3,4,5,6,7,8,10,11,12,13,14] for the Status element as a child of the Store element in the Search response.");
                #endregion

                #region Capture code for Range
                if (searchResponse.ResponseData.Response.Store.Range != null)
                {
                    Regex rangeRegex = new Regex("0-[0-9]*");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3725, the value for Range element is {0}", searchResponse.ResponseData.Response.Store.Range);

                    // Verify MS-ASCMD requirement: MS-ASCMD_R3725
                    Site.CaptureRequirementIfIsTrue(
                        rangeRegex.IsMatch(searchResponse.ResponseData.Response.Store.Range),
                        3725,
                        @"[In Range(Search)] The format of the Range element value is in the form of a zero-based index specifier, formed with a zero, a hyphen, and another numeric value: ""m-n.""");                    
                }
                #endregion

                #region Capture code for Result
                if (searchResponse.ResponseData.Response.Store.Status.Equals("1"))
                {
                    Site.Assert.IsTrue(searchResponse.ResponseData.Response.Store.Result != null && searchResponse.ResponseData.Response.Store.Result.Length > 0, "The Result element in Search command response should not be null.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3844");

                    // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        3844,
                        @"[In Result(Search)] The Result element is a required child element of the Store element in Search command responses that serves a container for an individual matching mailbox items.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2512");

                    // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2512,
                        @"[In Result(Search)] Element Result in Search command response (section 2.2.2.15), the parent element is Store (section 2.2.3.168.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2513");

                    // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2513,
                        @"[In Result(Search)] Element Result in Search command response (section 2.2.2.15), the child elements are airsync:Class (section 2.2.3.27.4), LongId (section 2.2.3.93.2), airsync:CollectionId (section 2.2.3.30.4), Properties (section 2.2.3.132.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2514");

                    // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2514,
                        @"[In Result(Search)] Element Result in Search command response (section 2.2.2.15), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2515");

                    // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        2515,
                        @"[In Result(Search)] Element Result in Search command response (section 2.2.2.15), the number allowed is 1...N (required).");

                    this.VerifyContainerDataType();

                    foreach (SearchResponseStoreResult result in searchResponse.ResponseData.Response.Store.Result)
                    {
                        #region Capture code for Class
                        if (result.Class != null)
                        {
                            Site.Assert.IsNotNull(result.Class, "The Class element should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5849");

                            // If the schema validation result is true and Class(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                5849,
                                @"[In Class(Search)] The airsync:Class element is a required child element of the Result element in Search command responses.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1319");

                            // If the schema validation result is true and Class(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1319,
                                @"[In Class(Search)] Element Class (Search) in Search command response, the parent element is Result (section 2.2.3.146.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1320");

                            // If the schema validation result is true and Class(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1320,
                                @"[In Class(Search)] None [Element Class (Search) in Search command response has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1321");

                            // If the schema validation result is true and Class(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1321,
                                @"[In Class(Search)] Element Class(Search) in Search command response, the data type is string.");

                            this.VerifyStringDataType();
                        }

                        #endregion

                        #region Capture code for LongId
                        if (result.LongId != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1836");

                            // If the schema validation result is true and LongId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1836,
                                @"[In LongId(Search)] Element LongId in Search command response (section 2.2.2.15), the parent element is Result (section 2.2.3.146.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1837");

                            // If the schema validation result is true and LongId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1837,
                                @"[In LongId(Search)] None [Element LongId in Search command response (section 2.2.2.15) has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1838");

                            // If the schema validation result is true and LongId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1838,
                                @"[In LongId(Search)] Element LongId in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1839");

                            // If the schema validation result is true and LongId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1839,
                                @"[In LongId(Search)] Element LongId in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                            this.VerifyStringDataType();
                        }
                        #endregion

                        #region Capture code for CollectionId
                        if (result.CollectionId != null)
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1392");

                            // If the schema validation result is true and CollectionId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1392,
                                @"[In CollectionId(Search)] Element CollectionId in Search command response,the parent element is Result (section 2.2.3.147.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1393");

                            // If the schema validation result is true and CollectionId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1393,
                                @"[In CollectionId(Search)] None [Element CollectionId in Search command response has no child element.]");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1394");

                            // If the schema validation result is true and CollectionId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1394,
                                @"[In CollectionId(Search)] Element CollectionId in Search command response, the data type is string.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1395");

                            // If the schema validation result is true and CollectionId(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1395,
                                @"[In CollectionId(Search)] Element CollectionId in Search command response, the number allowed is 0...N (optional).");

                            this.VerifyStringDataType();
                        }
                        #endregion

                        #region Capture code for Properties
                        if (result.Properties != null)
                        {
                            Site.Assert.IsTrue(result.Properties != null && result.Properties.Items.Length > 0, "The Properties element should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3687");

                            // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                3687,
                                @"[In Properties(Search)] The Properties element is a required child element of the Result element in Search command responses that contains the properties that are returned for item(s) in the response.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2402");

                            // If the schema validation result is true and Result(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2402,
                                @"[In Properties(Search)] Element Properties in Search command response (section 2.2.2.15), the parent element is Result (section 2.2.3.146.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2403");

                            // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2403,
                                @"[In Properties(Search)] Element Properties in Search command response (section 2.2.2.15), the child elements contain airsyncbase:Attachments ([MS-ASAIRS] section 2.2.2.8), airsyncbase:Body ([MS-ASAIRS] section 2.2.2.9), airsyncbase:BodyPart ([MS-ASAIRS] section 2.2.2.10), gal:Picture (section 2.2.3.129.2), rm:RightsManagementLicense ([MS-ASRM] section 2.2.2.14).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5751");

                            // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                5751,
                                @"[In Properties(Search)] Element Properties in Search command response (section 2.2.2.15), the child elements contain Data elements from the content classes.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2405");

                            // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2405,
                                @"[In Properties(Search)] Element Properties in Search command response (section 2.2.2.15), the data type is Container ([MS-ASDTYPE] section 2.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2406");

                            // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2406,
                                @"[In Properties(Search)] Element Properties in Search command response (section 2.2.2.15), the number allowed is 1...1 (required).");

                            this.VerifyContainerDataType();

                            if (result.Properties.ItemsElementName != null && result.Properties.ItemsElementName.Length > 0)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3848");

                                // If the schema validation result is true and the Properties element contains a list of properties, this requirement can be verified.
                                Site.CaptureRequirement(
                                    3848,
                                    @"[In Result(Search)] Inside the Result element, the Properties element contains a list of nonempty text properties on the entry.");

                                for (int j = 0; j < result.Properties.ItemsElementName.Length; j++)
                                {
                                    #region Capture code for Picture
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Picture1 && (SearchResponseStoreResultPropertiesPicture)result.Properties.Items[j] != null)
                                    {
                                        SearchResponseStoreResultPropertiesPicture picture = (SearchResponseStoreResultPropertiesPicture)result.Properties.Items[j];

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2380");

                                        // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2380,
                                            @"[In Picture(Search)] Element Picture in Search command response, the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2381");

                                        // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2381,
                                            @"[In Picture(Search)] Element Picture in Search command response, the child elements are Status (section 2.2.3.167.12) ,gal:Data (section 2.2.3.39.3).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2382");

                                        // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2382,
                                            @"[In Picture(Search)] Element Picture in Search command response, the data type is container.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2383");

                                        // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2383,
                                            @"[In Picture(Search)] Element Picture in Search command response, the number allowed is 0…1 (optional).");

                                        this.VerifyContainerDataType();

                                        #region Capture code for Status
                                        Site.Assert.IsNotNull(picture.Status, "The Status element in Picture element should not be null.");
                                        this.VerifyStatusElementForSearch();
                                        #endregion

                                        #region Capture code for Data
                                        if (picture.Data != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1470");

                                            // If the schema validation result is true and Data is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                1470,
                                                @"[In Data(Search)] Element Data in Search command response (section 2.2.2.15), the parent element is gal:Picture (section 2.2.3.129.2).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1471");

                                            // If the schema validation result is true and Data is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                1471,
                                                @"[In Data(Search)] None [ Element Data in Search command response (section 2.2.2.15) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1472");

                                            // If the schema validation result is true and Data is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                1472,
                                                @"[In Data(Search)] Element Data in Search command response (section 2.2.2.15), the data type is Byte array.");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1473");

                                            // If the schema validation result is true and Data is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                1473,
                                                @"[In Data(Search)] Element Data in Search command response (section 2.2.2.15), the number allowed is 0…1 (optional).");
                                        }
                                        #endregion
                                    }
                                    #endregion

                                    #region Capture code for Alias
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Alias1 && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5288");

                                        // If the schema validation result is true and Alias (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5288,
                                            @"[In Alias (Search)] Element Alias in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5289");

                                        // If the schema validation result is true and Alias (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5289,
                                            @"[In Alias (Search)] None [Element Alias in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5290");

                                        // If the schema validation result is true and Alias (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5290,
                                            @"[In Alias (Search)] Element Alias in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5291");

                                        // If the schema validation result is true and Alias (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5291,
                                            @"[In Alias (Search)] Element Alias in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for Company
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Company && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5295");

                                        // If the schema validation result is true and Company (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5295,
                                            @"[In Company (Search)] Element Company in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5296");

                                        // If the schema validation result is true and Company (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5296,
                                            @"[In Company (Search)] None [Element Company in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5297");

                                        // If the schema validation result is true and Company (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5297,
                                            @"[In Company (Search)] Element Company in Search command response (section2.2.2.15),the data type is string ([MS-ASDTYPE] section 2.6).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5298");

                                        // If the schema validation result is true and Company (Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5298,
                                            @"[In Company (Search)] Element Company in Search command response (section2.2.2.15),the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for DisplayName
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.DisplayName1 && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5302");

                                        // If the schema validation result is true and DisplayName(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5302,
                                            @"[In DisplayName(Search)] Element DisplayName in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5303");

                                        // If the schema validation result is true and DisplayName(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5303,
                                            @"[In DisplayName(Search)] None [Element DisplayName in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5304");

                                        // If the schema validation result is true and DisplayName(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5304,
                                            @"[In DisplayName(Search)] Element DisplayName in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5305");

                                        // If the schema validation result is true and DisplayName(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5305,
                                            @"[In DisplayName(Search)] Element DisplayName in Search command response (section 2.2.2.15), number allowed 0...1 (optional)");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for LinkId
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.LinkId && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1559");

                                        // If the schema validation result is true and LinkId is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1559,
                                            @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in Search command response, the parent element is search:Properties (section 2.2.3.128.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1560");

                                        // If the schema validation result is true and LinkId is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1560,
                                            @"[In documentlibrary:LinkId] None [Element documentlibrary:LinkId in Search command response has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1561");

                                        // If the schema validation result is true and LinkId is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1561,
                                            @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in Search command response, the data type is URI.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1562");

                                        // If the schema validation result is true and LinkId is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1562,
                                            @"[In documentlibrary:LinkId] Element documentlibrary:LinkId in Search command response, the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for EmailAddress
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.EmailAddress && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5318");

                                        // If the schema validation result is true and EmailAddress(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5318,
                                            @"[In EmailAddress(Search)] Element EmailAddress in Search command response (section 2.2.2.14), the parent element is Properties (section 2.2.3.128.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5319");

                                        // If the schema validation result is true and EmailAddress(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5319,
                                            @"[In EmailAddress(Search)] None [Element EmailAddress in Search command response (section 2.2.2.14) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5320");

                                        // If the schema validation result is true and EmailAddress(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5320,
                                            @"[In EmailAddress(Search)] Element EmailAddress in Search command response (section 2.2.2.14), the data type is string ([MS-ASDTYPE] section 2.6).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5321");

                                        // If the schema validation result is true and EmailAddress(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5321,
                                            @"[In EmailAddress(Search)] Element EmailAddress in Search command response (section 2.2.2.14), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for FirstName
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.FirstName1 && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5325");

                                        // If the schema validation result is true and FirstName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5325,
                                            @"[In FirstName] Element FirstName in Search command response (section 2.2.2.14), the parent element is Properties (section 2.2.3.128.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5326");

                                        // If the schema validation result is true and FirstName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5326,
                                            @"[In FirstName] None [Element FirstName in Search command response (section 2.2.2.14) has no child element]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5327");

                                        // If the schema validation result is true and FirstName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5327,
                                            @"[In FirstName] Element FirstName in Search command response (section 2.2.2.14), the data type is string ([MS-ASDTYPE] section 2.6).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5328");

                                        // If the schema validation result is true and FirstName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5328,
                                            @"[In FirstName] Element FirstName in Search command response (section 2.2.2.14), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for HomePhone
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.HomePhone && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5332");

                                        // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5332,
                                            @"[In HomePhone] Element HomePhone in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5333");

                                        // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5333,
                                            @"[In HomePhone] None [Element HomePhone in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5334");

                                        // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5334,
                                            @"[In HomePhone] Element HomePhone in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5335");

                                        // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5335,
                                            @"[In HomePhone] Element HomePhone in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for LastName
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.LastName1 && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5339");

                                        // If the schema validation result is true and LastName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5339,
                                            @"[In LastName] Element LastName in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5340");

                                        // If the schema validation result is true and LastName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5340,
                                            @"[In LastName] None [Element LastName in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5341");

                                        // If the schema validation result is true and LastName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5341,
                                            @"[In LastName] Element LastName in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5342");

                                        // If the schema validation result is true and LastName is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5342,
                                            @"[In LastName] Element LastName in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for MobilePhone
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.MobilePhone && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5355");

                                        // If the schema validation result is true and MobilePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5355,
                                            @"[In MobilePhone] Element MobilePhone in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5356");

                                        // If the schema validation result is true and MobilePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5356,
                                            @"[In MobilePhone] None [Element MobilePhone in Search command response (section 2.2.2.15) has no child element]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5357");

                                        // If the schema validation result is true and MobilePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5357,
                                            @"[In MobilePhone] Element MobilePhone in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5358");

                                        // If the schema validation result is true and MobilePhone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5358,
                                            @"[In MobilePhone] Element MobilePhone in Search command response (section 2.2.2.15), the number allowed is 0...1 (optional).");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for Office
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Office && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5362");

                                        // If the schema validation result is true and Office is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5362,
                                            @"[In Office] Element Office in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5363");

                                        // If the schema validation result is true and Office is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5363,
                                            @"[In Office] None [Element Office in Search command response (section 2.2.2.15) has no child element]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5364");

                                        // If the schema validation result is true and Office is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5364,
                                            @"[In Office] Element Office in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5365");

                                        // If the schema validation result is true and Office is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5365,
                                            @"[In Office] Element Office in Search command response (section 2.2.2.15), number allowed 0...1 (optional)");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for Phone
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Phone && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5370");

                                        // If the schema validation result is true and Phone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5370,
                                            @"[In Phone] Element Phone in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5371");

                                        // If the schema validation result is true and Phone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5371,
                                            @"[In Phone] None [Element Phone in Search command response (section 2.2.2.15) has no child element]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5372");

                                        // If the schema validation result is true and Phone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5372,
                                            @"[In Phone] Element Phone in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5373");

                                        // If the schema validation result is true and Phone is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5373,
                                            @"[In Phone] Element Phone in Search command response (section 2.2.2.15), number allowed 0...1 (optional)");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion

                                    #region Capture code for Title
                                    if (result.Properties.ItemsElementName[j] == ItemsChoiceType6.Title1 && (string)result.Properties.Items[j] != null)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5393");

                                        // If the schema validation result is true and Title is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5393,
                                            @"[In Title(Search)] Element Title in Search command response (section 2.2.2.15), the parent element is Properties (section 2.2.3.132.2)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5394");

                                        // If the schema validation result is true and Title is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5394,
                                            @"[In Title(Search)] None [Element Title in Search command response (section 2.2.2.15) has no child element.]");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5395");

                                        // If the schema validation result is true and Title is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5395,
                                            @"[In Title(Search)] Element Title in Search command response (section 2.2.2.15), the data type is string ([MS-ASDTYPE] section 2.7)");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5396");

                                        // If the schema validation result is true and Title is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            5396,
                                            @"[In Title(Search)] Element Title in Search command response (section 2.2.2.15), number allowed 0...1 (optional)");

                                        this.VerifyStringDataType();
                                    }
                                    #endregion
                                }
                            }
                        }

                        #endregion
                    }
                }
                #endregion
                #endregion
            }
            #endregion
        }
        #endregion

        #region Capture code for Find command
        /// <summary>
        /// This method is used to verify the Find response related requirements.
        /// </summary>
        /// <param name="findResponse">Find command response.</param>
        private void VerifyFindCommand(Microsoft.Protocols.TestSuites.Common.FindResponse findResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(findResponse.ResponseData, "The Find element should not be null.");

            #region Capture code for Find Store
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65591902");

            // If the schema validation result is true and Find is not null, this requirement can be verified.
            Site.CaptureRequirement(
                65591902,
                @"[In Find] The Find element is a required element in Find command responses that identifies the body of the HTTP POST as containing a Find command (section 2.2.3.69).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65591908");

            // If the schema validation result is true and Find is not null, this requirement can be verified.
            Site.CaptureRequirement(
                65591908,
                @"[In Find] None [Element Find in Find command response, has no Parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65591909");

            // If the schema validation result is true and Find is not null, this requirement can be verified.
            Site.CaptureRequirement(
                65591909,
                @"[In Find] Element Find in Find command response, the child elements are Status (section 2.2.3.177.2), Response (section 2.2.3.153.2)");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65591910");

            // If the schema validation result is true and Find is not null, this requirement can be verified.
            Site.CaptureRequirement(
                65591910,
                @"[In Find] Element Find in Find command response, the data type is container ([MS-ASDTYPE] section 2.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65591911");

            // If the schema validation result is true and Find is not null, this requirement can be verified.
            Site.CaptureRequirement(
                65591911,
                @"[In Find] Element Find in Find command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for Status
            Site.Assert.IsNotNull(findResponse.ResponseData.Status, "The Status element should not be null.");
            this.VerifyStatusElementForFind();            
            int status;
            Site.Assert.IsTrue(int.TryParse(findResponse.ResponseData.Status, out status), "The Status element should be an integer.");
            this.VerifyIntegerDataType();
            Common.VerifyActualValues("Status(Find)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4" }), findResponse.ResponseData.Status, this.Site);
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72172506");

            // If the schema validation result is true, this requirement can be verified.
            Site.CaptureRequirement(
                72172506,
                @"[In Status (Find)] The following table specifies valid values [1,2,3,4] for the Status element as a child of the Store element in the Search response.");

            #endregion

            #region Capture code for Response
            if (findResponse.ResponseData.Response != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R70481804");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    70481804,
                    @"[In Response (Find)] Element Response in Find command response (section 2.2.1.2), the parent element is Find (section 2.2.3.69).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R70481805");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    70481805,
                    @"[In Response (Find)] Element Response in Find command response (section 2.2.1.2), the child elements are itemoperations:Store (section 2.2.3.178.1) 
Status (section 2.2.3.177.2)
Result (section 2.2.3.155.1)
Range (section 2.2.3.143.1)
Total (section 2.2.3.184.1).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R70481806");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    70481806,
                    @"[In Response (Find)] Element Response in Find command response (section 2.2.1.2), the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R70481807");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    70481807,
                    @"[In Response (Find)] Element Response in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                #region Capture code for Store
               
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251801");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251801,
                    @"[In Store(Find)] The itemoperations:Store element is a required child element of the Response element in Find command responses that specifies the name of the store to which the parent operation applies.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251803");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251803,
                    @"[In Store(Find)] Element Store in , the parent element is Response (section 2.2.3.153.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251804");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251804,
                    @"[In Store(Find)] None [Element Store in  has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251805");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251805,
                    @"[In Store(Find)] Element Store in , the data type is string ([MS-ASDTYPE] section 2.7).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251806");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251806,
                    @"[In Store(Find)] Element Store in , the number allowed is 1...1 (required).");
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R45251807");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    45251807,
                    @"[In Store(Find)] In the Find command response, the value of the Store element will be 'Mailbox'.");
                #endregion                

                #region Capture code for Status
                Site.Assert.IsNotNull(findResponse.ResponseData.Response.Status, "The Status element should not be null.");

                this.VerifyStatusElementForFind();

                Common.VerifyActualValues("Status", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4"}), findResponse.ResponseData.Response.Status, this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R72172506");

                // Verify MS-ASCMD requirement: MS-ASCMD_R72172506
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    72172506,
                    @"[In Status (Find)] The following table specifies valid values [1,2,3,4] for the Status element as a child of the Store element in the Search response.");
                #endregion

                #region Capture code for Result
                if (findResponse.ResponseData.Response.Results != null)
                {
                    foreach (FindResponseResult result in findResponse.ResponseData.Response.Results)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38391804");

                        // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            38391804,
                            @"[In Result (Find)] One Result element is present for each match that is found.");

                        if (result != null)
                        {
                            #region Capture code for Result element.
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R99910");

                            // If the schema validation result is true, this requirement can be verified.
                            Site.CaptureRequirement(
                                99910,
                                @"[In In Result (Find)] Element Result in Find command response (section 2.2.1.2), the parent element is Search (section 2.2.3.153.2).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R99911");

                            // If the schema validation result is true, this requirement can be verified.
                            Site.CaptureRequirement(
                                99911,
                                @"[In In Result (Find)] Element Result in Find command response (section 2.2.1.2), the child element is airsync:Class (section 2.2.3.27.1),airsync:ServerId (section 2.2.3.166.1),airsync:CollectionId (section 2.2.3.30.1),Properties (section 2.2.3.139.1).");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R99912");

                            // If the schema validation result is true, this requirement can be verified.
                            Site.CaptureRequirement(
                                99912,
                                @"[In Result (Find)] Element Result in Find command response (section 2.2.1.2), the data type is container ([MS-ASDTYPE] section2.2 .");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R99913");

                            // If the schema validation result is true, this requirement can be verified.
                            Site.CaptureRequirement(
                                99913,
                                @"[In Result (Find)] Element Result in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                            this.VerifyContainerDataType();
                            #endregion

                            #region Capture code for Class                            
                            if (result.Class!= null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9010602");

                                // If the schema validation result is true and Class(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9010602,
                                    @"[In Class (Find)] The airsync:Class element is a required child element of a required child element of the Result element in Find command responses.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9010609");

                                // If the schema validation result is true and Class(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9010609,
                                    @"[In Class (Find)] Element Class(Find) in Find command response, the parent element is Result (section 2.2.3.155.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9010610");

                                // If the schema validation result is true and Class(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9010610,
                                    @"[In Class (Find)] Element Class(Find) in Find command response, has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9010611");

                                // If the schema validation result is true and Class(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9010611,
                                    @"[In Class (Find)] Element Class(Find) in Find command response, the data type is string.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9010612");

                                // If the schema validation result is true and Class(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9010612,
                                    @"[In Class (Find)] Element Class(Find) in Find command response, the number allowed is 1...1 (required).");

                                this.VerifyStringDataType();
                            }                            
                            #endregion

                            #region Capture code for CollectionId                            
                            if (result.CollectionId != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9820602");

                                // If the schema validation result is true and CollectionId(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9820602,
                                    @"[In CollectionId (Find)] The airsync:CollectionId element is a required child element of the Result element in Find command responses that specifies the folder in which the item was found.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9820608");

                                // If the schema validation result is true and CollectionId(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9820608,
                                    @"[In CollectionId (Find)] Element CollectionId in Find command response, the parent element is Result (section 2.2.3.155.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9820609");

                                // If the schema validation result is true and CollectionId(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9820609,
                                    @"[In CollectionId (Find)] None [Element in Find command response has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9820610");

                                // If the schema validation result is true and CollectionId(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9820610,
                                    @"[In CollectionId (Find)] Element CollectionId in Find command response, the data type is string.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R9820611");

                                // If the schema validation result is true and CollectionId(Find) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    9820611,
                                    @"[In CollectionId (Find)] Element CollectionId in Find command response, the number allowed is 1...1 (required).");

                                this.VerifyStringDataType();
                            }
                            #endregion

                            #region Capture code for ServerId                            
                            if (result.ServerId != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38971801");

                                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    38971801,
                                    @"[In ServerId (Find)] The airsync:ServerId element is a required child element of the Result element under the Find element (section 2.2.3.69) in Find command responses.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38971803");

                                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    38971803,
                                    @"[In ServerId (Find)] Element ServerId in Find command response (section 2.2.1.2), the parent element is Result (section 2.2.3.155.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38971804");

                                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    38971804,
                                    @"[In ServerId (Find)] None [Element ServerId in Find command response (section 2.2.1.2), has no child element.].");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38971805");

                                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    38971805,
                                    @"[In ServerId (Find)] Element ServerId in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R38971806");

                                // If the schema validation result is true and ServerId(FolderCreate) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    38971806,
                                    @"[In ServerId (Find)] Element ServerId in Find command response (section 2.2.1.2), the number allowed is 1…1 (required).");

                                this.VerifyStringDataType();
                            }
                            #endregion

                            #region Capture code for Properties                            
                            if (result.Properties != null)
                            {
                                Site.Assert.IsTrue(result.Properties != null && result.Properties.Items.Length > 0, "The Properties element should not be null.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36841701");

                                // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    36841701,
                                    @"[In Properties (Find)] The Properties element is a required child element of the Result element in Find command responses that contains the properties that are returned for an item in the response.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36841703");

                                // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    36841703,
                                    @"[In Properties (Find)] Element Properties in Find command response (section 2.2.1.2), the parent element is Result (section 2.2.3.155.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36841704");

                                // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    36841704,
                                    @"[In Properties (Find)] Element Properties in Find command response (section 2.2.1.2)  the child elements are Subject ([MS-ASEMAIL] section 2.2.2.75.1), 
                                    DateReceived ([MS-ASEMAIL] section 2.2.2.24), DisplayTo ([MS-ASEMAIL] section 2.2.2.29), 
                                    DisplayCc (section 2.2.3.48)
                                    DisplayBcc (section 2.2.3.47)
                                    Importance ([MS-ASEMAIL] section 2.2.2.38)
                                    Read ([MS-ASEMAIL] section 2.2.2.58)
                                    IsDraft ([MS-ASEMAIL] section 2.2.2.42)
                                    Preview (section 2.2.3.137)
                                    HasAttachments (section 2.2.3.87)
                                    From ([MS-ASEMAIL] section 2.2.2.36)
                                    gal:DisplayName (section 2.2.3.49.2)
                                    gal:Phone (section 2.2.3.133.1)
                                    gal:Office (section 2.2.3.121.1)
                                    gal:Title (section 2.2.3.182.1)
                                    gal:Company (section 2.2.3.33.1)
                                    gal:Alias (section 2.2.3.9)
                                    gal:FirstName (section 2.2.3.70)
                                    gal:LastName (section 2.2.3.95)
                                    gal:HomePhone (section 2.2.3.89)
                                    gal:MobilePhone (section 2.2.3.114)
                                    gal:EmailAddress (section 2.2.3.55.1)
                                    gal:Picture (section 2.2.3.135.1).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36841705");

                                // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    36841705,
                                    @"[In Properties (Find)] Element Properties in Find command response (section 2.2.1.2), the data type is Container ([MS-ASDTYPE] section 2.2).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36841706");

                                // If the schema validation result is true and Properties(Search) is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    36841706,
                                    @"[In Properties (Find)] Element Properties in Find command response (section 2.2.1.2), the number allowed is 1...1 (required).");

                                if (result.Properties.ItemsElementName != null && result.Properties.ItemsElementName.Length > 0)
                                {
                                    for (int j = 0; j < result.Properties.ItemsElementName.Length; j++)
                                    {
                                        #region Capture code for Picture
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Picture && (FindResponseResult)result.Properties.Items[j] != null)
                                        {
                                            SearchResponseStoreResultPropertiesPicture picture = (SearchResponseStoreResultPropertiesPicture)result.Properties.Items[j];
                                            #region Capture code for Status
                                            Site.Assert.IsNotNull(picture.Status, "The Status element in Picture element should not be null.");
                                            this.VerifyStatusElementForSearch();
                                            #endregion
                                            // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36691708");

                                        // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            36691708,
                                            @"[In Picture (Find)] Element Picture in Find command response, the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36691709");

                                            // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                36691709,
                                                @"[In Picture (Find)] Element Picture in Find command response, the child elements are Status (section 2.2.3.177.2), gal:Data (section 2.2.3.39.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36691710");

                                            // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                36691710,
                                                @"[In Picture (Find)] Element Picture in Find command response, the data type is container.");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R36691711");

                                            // If the schema validation result is true and Picture(Search) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                36691711,
                                                @"[In Picture (Find)] Element Picture in Find command response, the number allowed is 0…1 (optional).");
                                            #region Capture code for Data
                                            if (picture.Data != null)
                                            {
                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R21270702");

                                                // If the schema validation result is true and Data is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    21270702,
                                                    @"[In Data (Find)] Element Data in Find command response (section 2.2.1.16), the parent element is gal:Picture (section 2.2.3.135.1).");

                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R21270703");

                                                // If the schema validation result is true and Data is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    21270703,
                                                    @"[In Data (Find)] None [Element Data in Find command response (section 2.2.1.16) has no child element.]");

                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R21270704");

                                                // If the schema validation result is true and Data is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    21270704,
                                                    @"[In Data (Find)] Element Data in Find command response (section 2.2.1.16), the data type is Byte array.");

                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R21270705");

                                                // If the schema validation result is true and Data is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    21270705,
                                                    @"[In Data (Find)] Element Data in Find command response (section 2.2.1.16), the number allowed is 0…1 (optional).");
                                            }
                                            #endregion
                                        }
                                        #endregion

                                        #region  Capture code for Alias
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Alias && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R61641063");

                                            // If the schema validation result is true and Alias(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                61641063,
                                                @"[In Alias (Find)] Element Alias in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R61641064");

                                            // If the schema validation result is true and Alias(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                61641064,
                                                @"[In Alias (Find)] None [Element Alias in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R61641065");

                                            // If the schema validation result is true and Alias(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                61641065,
                                                @"[In Alias (Find)] Element Alias in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R61641066");

                                            // If the schema validation result is true and Alias(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                61641066,
                                                @"[In Alias (Find)] Element Alias in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code Company
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Company && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R63510607");

                                            // If the schema validation result is true and Company (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                63510607,
                                                @"[In Company (Find)] Element Company (Find) in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R63510608");

                                            // If the schema validation result is true and Company (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                63510608,
                                                @"[In Company (Find)] None [Element Company (Find) in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R63510609");

                                            // If the schema validation result is true and Company (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                63510609,
                                                @"[In Company (Find)] Element Company (Find) in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R63510610");

                                            // If the schema validation result is true and Company (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                63510610,
                                                @"[In Company (Find)] Element Company (Find) in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for DisplayBcc
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.DisplayBcc && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390706");

                                            // If the schema validation result is true and DisplayBcc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390706,
                                                @"[In DisplayBcc] Element DisplayBcc in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390707");

                                            // If the schema validation result is true and DisplayBcc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390707,
                                                @"[In DisplayBcc] None [Element DisplayBcc in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390708");

                                            // If the schema validation result is true and DisplayBcc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390708,
                                                @"[In DisplayBcc] Element DisplayBcc in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390709");

                                            // If the schema validation result is true and DisplayBcc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390709,
                                                @"[In DisplayBcc] Element DisplayBcc in Find command response (section 2.2.1.2), the number allowed is 0…1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for DisplayCc
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.DisplayCc && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390719");

                                            // If the schema validation result is true and DisplayCc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390719,
                                                @"[In DisplayCc] Element DisplayCc in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390720");

                                            // If the schema validation result is true and DisplayCc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390720,
                                                @"[In DisplayCc] None [Element DisplayCc in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390721");

                                            // If the schema validation result is true and DisplayCc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390721,
                                                @"[In DisplayCc] Element DisplayCc in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64390722");

                                            // If the schema validation result is true and DisplayCc is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64390722,
                                                @"[In DisplayCc] Element DisplayCc in Find command response (section 2.2.1.2), the number allowed is 0…1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for DisplayName
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.DisplayName && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64430704");

                                            // If the schema validation result is true and DisplayName(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64430704,
                                                @"[In DisplayName (Find)] Element DisplayName in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64430705");

                                            // If the schema validation result is true and DisplayName(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64430705,
                                                @"[In DisplayName (Find)] None [Element DisplayName in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64430706");

                                            // If the schema validation result is true and DisplayName(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64430706,
                                                @"[In DisplayName (Find)] Element DisplayName in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R64430707");

                                            // If the schema validation result is true and DisplayName(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                64430707,
                                                @"[In DisplayName (Find)] Element DisplayName in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for EmailAddress
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.EmailAddress && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R53080804");

                                            // If the schema validation result is true and EmailAddress(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                53080804,
                                                @"[EmailAddress (Find)] Element EmailAddress in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R53080805");

                                            // If the schema validation result is true and EmailAddress(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                53080805,
                                                @"[EmailAddress (Find)] None [Element EmailAddress in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R53080806");

                                            // If the schema validation result is true and EmailAddress(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                53080806,
                                                @"[EmailAddress (Find)] Element EmailAddress in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R53080807");

                                            // If the schema validation result is true and EmailAddress(Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                53080807,
                                                @"[EmailAddress (Find)] Element EmailAddress in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for FirstName
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.FirstName && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65590808");

                                            // If the schema validation result is true and FirstName (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                65590808,
                                                @"[In FirstName (Find)] Element FirstName in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1) .");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65590809");

                                            // If the schema validation result is true and FirstName (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                65590809,
                                                @"[In FirstName (Find)] None [Element FirstName in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65590810");

                                            // If the schema validation result is true and FirstName (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                65590810,
                                                @"[In FirstName (Find)] Element FirstName in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R65590811");

                                            // If the schema validation result is true and FirstName (Find) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                65590811,
                                                @"[In FirstName (Find)] Element FirstName in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion

                                        #region Capture code for HasAttachments
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.HasAttachments && (bool?)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66851603");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66851603,
                                                @"[In HasAttachments] Element HasAttachments in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66851604");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66851604,
                                                @"[In HasAttachments] None [Element HasAttachments in Find command response (section 2.2.1.2) has no child element.");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66851605");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66851605,
                                                @"[In HasAttachments] Element HasAttachments in Find command response (section 2.2.1.2), the data type is boolean ([MS-ASDTYPE] section 2.1.");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66851606");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66851606,
                                                @"[In HasAttachments] Element HasAttachments in Find command response (section 2.2.1.2), the number allowed is 0…1 (optional)");
                                        }
                                        #endregion

                                        #region Capture code for HomePhone
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.HomePhone && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66941607");

                                            // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66941607,
                                                @"[In HomePhone(Find)] Element HomePhone in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66941608");

                                            // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66941608,
                                                @"[In HomePhone(Find)] None [Element HomePhone in Find command response (section 2.2.1.2) has no child element.].");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66941609");

                                            // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66941609,
                                                @"[In HomePhone(Find)] Element HomePhone in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R66941610");

                                            // If the schema validation result is true and HomePhone is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                66941610,
                                                @"[In HomePhone(Find)] Element HomePhone in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                        }
                                        #endregion

                                        #region Capture code for LastName
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.LastName && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R67291607");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                67291607,
                                                @"[LastName (Find)] Element LastName in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R67291608");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                67291608,
                                                @"[LastName (Find)] None [Element in Find command response (section 2.2.1.2) has no child element.].");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R67291609");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                67291609,
                                                @"[LastName (Find)] Element LastName in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R67291610");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                67291610,
                                                @"[LastName (Find)] Element LastName in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                        }
                                        #endregion

                                        #region Capture code for MobilePhone
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.MobilePhone && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68291608");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68291608,
                                                @"[In MobilePhone (Find)] Element MobilePhone in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68291610");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68291610,
                                                @"[In MobilePhone (Find)] None [Element MobilePhone in Find command response (section 2.2.1.2) has no child element.].");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68291611");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68291611,
                                                @"[In MobilePhone (Find)] Element MobilePhone  in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68291612");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68291612,
                                                @"[In MobilePhone (Find)] Element MobilePhone in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");
                                        }
                                        #endregion

                                        #region Capture code for Office
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Office && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651724");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68651724,
                                                @"[In Office (Find)] Element Office in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651725");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68651725,
                                                @"[In Office (Find)] None [Element in Find command response (section 2.2.1.2) has no child element.]).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651726");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68651726,
                                                @"[In Office (Find)] Element Office in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R68651727");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                68651727,
                                                @"[In Office (Find)] Element Office in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");
                                        }
                                        #endregion

                                        #region Capture code for Phone
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Phone && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69581708");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69581708,
                                                @"[In Phone (Find)] Element Phone in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69581709");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69581709,
                                                @"[In Phone (Find)] None [Element Phone in Find command response (section 2.2.1.2) has no child element.");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69581710");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69581710,
                                                @"[In Phone (Find)] Element Phone in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69581711");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69581711,
                                                @"[In Phone (Find)] Element Phone in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");
                                        }
                                        #endregion

                                        #region Capture code for Preview
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Preview && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69781704");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69781704,
                                                @"[In Preview] Element in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69781705");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69781705,
                                                @"[In Preview] None [Element Preview in Find command response (section 2.2.1.2) has no child element.].");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69781706");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69781706,
                                                @"[In Preview] Element Preview in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R69781707");

                                            // If the schema validation result is true and Certificates(ResolveRecipients) is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                69781707,
                                                @"[In Preview] Element Preview in Find command response (section 2.2.1.2), the number allowed is 0…1 (optional).");
                                        
                                    }
                                        #endregion

                                        #region Capture code for Title
                                        if (result.Properties.ItemsElementName[j] == ItemsChoiceType14.Title && (string)result.Properties.Items[j] != null)
                                        {
                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R73291807");

                                            // If the schema validation result is true and Title is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                73291807,
                                                @"[In Title(Find)] Element Title in Find command response (section 2.2.1.2), the parent element is Properties (section 2.2.3.139.1).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R73291808");

                                            // If the schema validation result is true and Title is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                73291808,
                                                @"[In Title(Find)] None [Element Title in Find command response (section 2.2.1.2) has no child element.]");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R73291809");

                                            // If the schema validation result is true and Title is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                                73291809,
                                                @"[In Title(Find)] Element Title in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                                            // Add the debug information.
                                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R73291810");

                                            // If the schema validation result is true and Title is not null, this requirement can be verified.
                                            Site.CaptureRequirement(
                                               73291810,
                                                @"[In Title(Find)]  Element Title in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");

                                            this.VerifyStringDataType();
                                        }
                                        #endregion
                                    }
                                }
                            }
                            #endregion
                        }
                    }
                }

                #endregion

                #region Capture code for Range                
                if (findResponse.ResponseData.Response.Range != null)
                {
                    //Regex rangeRegex = new Regex("1-[9]*");
                    Regex rangeRegex = new Regex("0-[0-9]*");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R37071811, the value for Range element is {0}", findResponse.ResponseData.Response.Range);

                    // Verify MS-ASCMD requirement: MS-ASCMD_R37071811
                    Site.CaptureRequirementIfIsTrue(
                        rangeRegex.IsMatch(findResponse.ResponseData.Response.Range),
                        37071811,
                        @"[In Range (Find)] The format of the Range element value is in the form of a zero-based index specifier, formed with a nonnegative integer, a hyphen, and another nonnegative integer of higher value than the first.");

                    uint m=uint.Parse(findResponse.ResponseData.Response.Range.Split('-')[0]);
                    uint n = uint.Parse(findResponse.ResponseData.Response.Range.Split('-')[1]);
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R37071812");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R37071812
                    Site.CaptureRequirementIfIsTrue(
                        m>=0&&m<=n&&n<=999,
                        37071812,
                        @"[In Range (Find)] The m indicates the lowest index of a zero-based array that would hold the items. The n indicates the highest index of a zero-based array that would hold the items. The Range element has possible values for m and n of 0 ≤ m ≤ n ≤ 999.");
                }
                #endregion

                #region Capture code for Total
                // Add the debug information.

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R46301803");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    46301803,
                    @"[In Total(Find)] Element Total in Find command response (section 2.2.1.2), the parent element is Response (section 2.2.3.153.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R46301804");

                // If the schema validation result is true, this requirement can be verified.
                Site.CaptureRequirement(
                    46301804,
                    @"[In Total(Find)] None [Element Total in Find command response (section 2.2.1.2) has no child element.]");
                int total;
                if (findResponse.ResponseData.Response.Store != null)
                {
                    Site.Assert.IsNotNull(findResponse.ResponseData.Response.Store, "The Store element in Find command response should not be null.");
                    
                }

                this.VerifyStringDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R46301805");

                // Verify MS-ASCMD requirement: MS-ASCMD_R2857
                Site.CaptureRequirement(
                    46301805,
                    @"[In Total(Find)] Element Total in Find command response (section 2.2.1.2), the data type is string ([MS-ASDTYPE] section 2.7).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R46301806");

                // If the schema validation result is true and Total(find) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    46301806,
                    @"[In Total(Find)] Element Total in Find command response (section 2.2.1.2), the number allowed is 0...1 (optional).");
                #endregion
            }
            #endregion
        }

        #region Capture code for SendMail command
        /// <summary>
        /// This method is used to verify the SendMail response related requirements.
        /// </summary>
        /// <param name="sendMailResponse">SendMail command response.</param>
        private void VerifySendMailCommand(SendMailResponse sendMailResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(sendMailResponse.ResponseData, "The SendMail element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3887");

            // If the schema validation result is true and SendMail is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3887,
                @"[In SendMail] The SendMail element is a required element in SendMail command requests and responses that identifies the body of the HTTP POST as containing a SendMail command (section 2.2.2.16).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2557");

            // If the schema validation result is true and SendMail is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2557,
                @"[In SendMail] None [Element SendMail in SendMail command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2558");

            // If the schema validation result is true and SendMail is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2558,
                @"[In SendMail] Element SendMail in SendMail command response, the child element is Status (section 2.2.3.167.13).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2559");

            // If the schema validation result is true and SendMail is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2559,
                @"[In SendMail] Element SendMail in SendMail command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2560");

            // If the schema validation result is true and SendMail is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2560,
                @"[In SendMail] Element SendMail in SendMail command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();

            if (sendMailResponse.ResponseData.Status != null)
            {
                int status;

                Site.Assert.IsTrue(int.TryParse(sendMailResponse.ResponseData.Status, out status), "The Status element in SendMail response should be an integer.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2751");

                // If the schema validation result is true and Status(SendMail) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2751,
                    @"[In Status(SendMail)] Element Status in SendMail command response,the parent element is SendMail (section 2.2.3.152).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2752");

                // If the schema validation result is true and Status(SendMail) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2752,
                    @"[In Status(SendMail)] None [Element Status in SendMail command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2753");

                // If the schema validation result is true, Status(SendMail) is an integer, this requirement can be verified.
                Site.CaptureRequirement(
                    2753,
                    @"[In Status(SendMail)] Element Status in SendMail command response, the data type is integer ([MS-ASDTYPE] section 2.6).");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2754");

                // If the schema validation result is true and Status(SendMail) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2754,
                    @"[In Status(SendMail)] Element Status in SendMail command response, the number allowed is 0...1 (optional).");
            }
        }
        #endregion

        #region Capture code for Settings command
        /// <summary>
        /// This method is used to verify the Settings response related requirements.
        /// </summary>
        /// <param name="settingsResponse">Settings command response.</param>
        private void VerifySettingsCommand(SettingsResponse settingsResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(settingsResponse.ResponseData, "The Settings element should not be null.");
            Site.Assert.IsNotNull(settingsResponse.ResponseData.Status, "The Settings status code should not be null");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R485");

            // If the schema validation result is true and Status(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                485,
                @"[In Settings] All property responses, regardless of the property, MUST contain a Status element to indicate success or failure.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R486");

            // If the schema validation result is true and Status(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                486,
                @"[In Settings] This Status node MUST be the first node in the property response.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3953");

            // If the schema validation result is true and Settings(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3953,
                @"[In Settings(Settings)] The Settings element is a required element in Settings command requests and responses that identifies the body of the HTTP POST as containing a Settings command (section 2.2.2.17).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2631");

            // If the schema validation result is true and Settings(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2631,
                @"[In Settings(Settings)] None [Element Settings in Settings command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2632");

            // If the schema validation result is true and Settings(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2632,
                @"[In Settings(Settings)] Element Settings in Settings command response, the child elements are RightsManagementInformation (section 2.2.3.147), Oof
,DeviceInformation, DevicePassword, UserInformation, Status (section 2.2.3.167.14).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2633");

            // If the schema validation result is true and Settings(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2633,
                @"[In Settings(Settings)] Element Settings in Settings command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2634");

            // If the schema validation result is true and Settings(Settings) is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2634,
                @"[In Settings(Settings)] Element Settings in Settings command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();

            #region Capture code for RightsManagementInformation
            if (settingsResponse.ResponseData.RightsManagementInformation != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2520");

                // If the schema validation result is true and RightsManagementInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2520,
                    @"[In RightsManagementInformation] Element RightsManagementInformation in Settings command response, the parent element is Settings.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2521");

                // If the schema validation result is true and RightsManagementInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2521,
                    @"[In RightsManagementInformation] Element RightsManagementInformation in Settings command response, the child elements are Get, Status (section 2.2.3.167.14).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2522");

                // If the schema validation result is true and RightsManagementInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2522,
                    @"[In RightsManagementInformation] Element RightsManagementInformation in Settings command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2523");

                // If the schema validation result is true and RightsManagementInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2523,
                    @"[In RightsManagementInformation] Element RightsManagementInformation in Settings command response, the number allowed is 0…1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Status
                Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Status, "As a child element of RightsManagementInformation, the Status element should not be null.");

                int status;

                Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.RightsManagementInformation.Status, out status), "As a child element of RightsManagementInformation, the Status element should be an integer.");

                this.VerifyStatusElementForSettings();

                Common.VerifyActualValues("Status(Settings)", AdapterHelper.ValidStatus(new string[] { "1", "2", "5", "6" }), settingsResponse.ResponseData.RightsManagementInformation.Status, this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4391");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4391
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4391,
                    @"[In Status(Settings)] The following table lists the valid values [1,2,5,6] for Status in a Settings command RightsManagementInformation Get operation, Oof Get operation, Oof Set operation, DeviceInformation Set operation, or UserInformation Get operation.");
                #endregion

                #region Capture code for Get
                Site.Assert.IsNotNull(settingsResponse.ResponseData.RightsManagementInformation.Get, "As a child element of RightsManagementInformation, the Get element should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3113");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    3113,
                    @"[In Get] The Get element is a required child element of the RightsManagementInformation element in Settings command RightsManagementInformation requests and responses.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1739");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1739,
                    @"[In Get] Element Get in Settings command RightsManagementInformation response, the parent element is  RightsManagementInformation.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1740");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1740,
                    @"[In Get] Element Get in Settings command RightsManagementInformation response, the child element is  rm:RightsManagementTemplates ([MS-ASRM] section 2.2.2.17).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1741");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1741,
                    @"[In Get] Element Get in Settings command RightsManagementInformation response, the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1742");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1742,
                    @"[In Get] Element Get in Settings command RightsManagementInformation response, the number allowed is  1…1 (required).");

                this.VerifyContainerDataType();
                #endregion
            }
            #endregion

            #region Capture code for Oof
            if (settingsResponse.ResponseData.Oof != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1991");

                // If the schema validation result is true and Oof is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1991,
                    @"[In Oof] Element Oof in Settings command response, the parent element is Settings.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1992");

                // If the schema validation result is true and Oof is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1992,
                    @"[In Oof] Element Oof in Settings command response, the child elements are Get, Status (section 2.2.3.167.14).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1993");

                // If the schema validation result is true and Oof is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1993,
                    @"[In Oof] Element Oof in Settings command response,  the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1994");

                // If the schema validation result is true and Oof is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1994,
                    @"[In Oof] Element Oof in Settings command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Status
                Site.Assert.IsNotNull(settingsResponse.ResponseData.Oof.Status, "As a child element of Oof, the Status element should not be null.");

                int status;

                Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.Oof.Status, out status), "As a child element of Oof, the Status element should be an integer.");

                this.VerifyStatusElementForSettings();
                #endregion

                #region Capture code for Get
                if (settingsResponse.ResponseData.Oof.Get != null)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1747");

                    // If the schema validation result is true and Get is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1747,
                        @"[In Get] Element Get in Settings command Oof response, the parent element is Oof.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1748");

                    // If the schema validation result is true and Get is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1748,
                        @"[In Get] Element Get in Settings command Oof response, the child elements are OofState (section 2.2.3.118), StartTime (section 2.2.3.166.2), EndTime (section 2.2.3.58.2), OofMessage (section 2.2.3.117).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1749");

                    // If the schema validation result is true and Get is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1749,
                        @"[In Get] Element Get in Settings command Oof response, the data type is container.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1750");

                    // If the schema validation result is true and Get is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1750,
                        @"[In Get] Element Get in Settings command Oof response, the number allowed is 0...1 (optional).");

                    this.VerifyContainerDataType();

                    #region Capture code for OofState
                    if (settingsResponse.ResponseData.Oof.Get.OofStateSpecified)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2007");

                        // If the schema validation result is true and OofState is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2007,
                            @"[In OofState] Element OofState in Settings command Oof response, the parent element is Get (section 2.2.3.79).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2008");

                        // If the schema validation result is true and OofState is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2008,
                            @"[In OofState] None [Element OofState in Settings command Oof response has no child element.]");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2009");

                        // If the schema validation result is true and OofState is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2009,
                            @"[In OofState] Element OofState in Settings command Oof response, the data type is integer.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2010");

                        // If the schema validation result is true and OofState is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2010,
                            @"[In OofState] Element OofState in Settings command Oof response, the number allowed is 0...1 (optional).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3530, the value for OofStatus element is {0}", settingsResponse.ResponseData.Oof.Get.OofState);

                        // Verify MS-ASCMD requirement: MS-ASCMD_R3530
                        Site.CaptureRequirementIfIsTrue(
                            settingsResponse.ResponseData.Oof.Get.OofState == OofState.Item0 || settingsResponse.ResponseData.Oof.Get.OofState == OofState.Item1 || settingsResponse.ResponseData.Oof.Get.OofState == OofState.Item2,
                            3530,
                            @"[In OofState] The following table lists the valid values [0, 1, 2] for OofState.");
                    }
                    #endregion

                    #region Capture code for StartTime

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2687");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        2687,
                        @"[In StartTime(Settings)] Element StartTime in Settings command Oof response, the parent element is Get (section 2.2.3.79).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2688");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        2688,
                        @"[In StartTime(Settings)] None [Element StartTime in Settings command Oof response has no child element.]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2689");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        2689,
                        @"[In StartTime(Settings)] Element StartTime in Settings command Oof response, the data type is datetime.");

                    this.VerifyDateTimeStructure();

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2690");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        2690,
                        @"[In StartTime(Settings)] Element StartTime in Settings command Oof response, the number allowed is 0...1 (optional).");

                    #endregion

                    #region Capture code for EndTime
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1623");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        1623,
                        @"[In EndTime(Settings)] Element EndTime in Settings command Oof response, the parent element is Get (section 2.2.3.75).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1624");

                    // If the schema validation result is true, this requirement can be verified.
                    Site.CaptureRequirement(
                        1624,
                        @"[In EndTime(Settings)] None [Element EndTime in Settings command Oof response has no child element .]");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1625");

                    // Verify MS-ASCMD requirement: MS-ASCMD_R1625
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(DateTime),
                        settingsResponse.ResponseData.Oof.Get.EndTime.GetType(),
                        1625,
                        @"[In EndTime(Settings)] Element EndTime in Settings command Oof response, the data type is datetime.");

                    this.VerifyDateTimeStructure();

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1626");

                    // If the schema validation result is true and EndTime(Settings) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1626,
                        @"[In EndTime(Settings)] Element EndTime in Settings command Oof response, the number allowed is 0...1 (optional).");
                    #endregion

                    #region Capture code for OofMessage
                    if (settingsResponse.ResponseData.Oof.Get.OofMessage != null && settingsResponse.ResponseData.Oof.Get.OofMessage.Length > 0)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1999");

                        // If the schema validation result is true and OofMessage is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1999,
                            @"[In OofMessage] Element OofMessage in Settings command Oof response, the parent element is Get (section 2.2.3.79).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2000");

                        // If the schema validation result is true and OofMessage is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2000,
                            @"[In OofMessage] Element OofMessage in Settings command Oof response, the child elements are AppliesToInternal, AppliesToExternalKnown, AppliesToExternalUnknown, Enabled, ReplyMessage, BodyType.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2001");

                        // If the schema validation result is true and OofMessage is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            2001,
                            @"[In OofMessage] Element OofMessage in Settings command Oof response, the data type is container.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2002");

                        // Verify MS-ASCMD requirement: MS-ASCMD_R2002
                        Site.CaptureRequirementIfIsTrue(
                            settingsResponse.ResponseData.Oof.Get.OofMessage.Length <= 3,
                            2002,
                            @"[In OofMessage] Element OofMessage in Settings command Oof response, the number allowed is 0...3 (optional).");

                        this.VerifyContainerDataType();

                        foreach (OofMessage oofMessage in settingsResponse.ResponseData.Oof.Get.OofMessage)
                        {
                            #region Capture code for Enabled
                            if (oofMessage.Enabled != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1607");

                                // If the schema validation result is true and Enabled is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1607,
                                    @"[In Enabled] Element Enabled in Settings command Oof request and response (section 2.2.2.16), the parent element is OofMessage (section 2.2.3.113).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1608");

                                // If the schema validation result is true and Enabled is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1608,
                                    @"[In Enabled] None [Element Enabled in Settings command Oof request and response (section 2.2.2.16) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1609");

                                // If the schema validation result is true and Enabled is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1609,
                                    @"[In Enabled] Element Enabled in Settings command Oof request and response (section 2.2.2.16), the data type is string ([MS-ASDTYPE] section 2.6).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1610");

                                // If the schema validation result is true and Enabled is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1610,
                                    @"[In Enabled] Element Enabled in Settings command Oof request and response (section 2.2.2.16), the number allowed is 0...1 (optional).");

                                this.VerifyStringDataType();
                            }
                            #endregion

                            #region Capture code for BodyType
                            if (oofMessage.BodyType != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1103");

                                // If the schema validation result is true and BodyType is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1103,
                                    @"[In BodyType] Element BodyType in Settings command Oof response , the parent element is OofMessage.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1104");

                                // If the schema validation result is true and BodyType is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1104,
                                    @"[In BodyType] None [Element BodyType in Settings command Oof response has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1105");

                                // If the schema validation result is true and BodyType is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1105,
                                    @"[In BodyType] Element BodyType in Settings command Oof response, the data type is string.");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R821, the value for BodyType element is {0}", oofMessage.BodyType);

                                // Verify MS-ASCMD requirement: MS-ASCMD_R821
                                Site.CaptureRequirementIfIsTrue(
                                    oofMessage.BodyType.Equals("Text", StringComparison.CurrentCultureIgnoreCase) || oofMessage.BodyType.Equals("HTML", StringComparison.CurrentCultureIgnoreCase),
                                    821,
                                    @"[In BodyType] The following are the permitted values for the BodyType element: Text, HTML.");
                            }
                            #endregion

                            #region Capture code for ReplyMessage
                            if (oofMessage.ReplyMessage != null)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2451");

                                // If the schema validation result is true and ReplyMessage is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2451,
                                    @"[In ReplyMessage] Element ReplyMessage in Settings command Oof request and response (section 2.2.2.17), the parent element is OofMessage (section 2.2.3.117).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2452");

                                // If the schema validation result is true and ReplyMessage is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2452,
                                    @"[In ReplyMessage] None [ Element ReplyMessage in Settings command Oof request and response (section 2.2.2.17) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2453");

                                // If the schema validation result is true and ReplyMessage is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2453,
                                    @"[In ReplyMessage] Element ReplyMessage in Settings command Oof request and response (section 2.2.2.17), the data type is string ([MS-ASDTYPE] section 2.7).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2454");

                                // If the schema validation result is true and ReplyMessage is not null, this requirement can be verified.
                                Site.CaptureRequirement(
                                    2454,
                                    @"[In ReplyMessage] Element ReplyMessage in Settings command Oof request and response (section 2.2.2.17), the number allowed is 0...1 (optional).");
                            }
                            #endregion Capture code for ReplyMessage
                        }

                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.LoadXml(settingsResponse.ResponseDataXML);

                        if (xmlDoc.DocumentElement.HasChildNodes)
                        {
                            XmlNodeList appliesToInternalNodes = xmlDoc.DocumentElement.GetElementsByTagName("AppliesToInternal");
                            XmlNodeList appliesToExternalKnownNodes = xmlDoc.DocumentElement.GetElementsByTagName("AppliesToExternalKnown");
                            XmlNodeList appliesToExternalUnknownNodes = xmlDoc.DocumentElement.GetElementsByTagName("AppliesToExternalUnknown");

                            #region Capture code for AppliesToInternal
                            if (appliesToInternalNodes.Count > 0)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1075");

                                // If the schema validation result is true and AppliesToInternal exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1075,
                                    @"[In AppliesToInternal] Element AppliesToInternal in Settings command Oof request and response (section 2.2.2.17),the parent elements is OofMessage (section 2.2.3.117).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1076");

                                // If the schema validation result is true and AppliesToInternal exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1076,
                                    @"[In AppliesToInternal] None [Element AppliesToInternal in Settings command Oof request and response (section 2.2.2.17) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1077");

                                // If the schema validation result is true and AppliesToInternal exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1077,
                                    @"[In AppliesToInternal] None [Element AppliesToInternal in Settings command Oof request and response (section 2.2.2.17), the data type is None.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1078");

                                // If the schema validation result is true and AppliesToInternal exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1078,
                                    @"[In AppliesToInternal] Element AppliesToInternal in Settings command Oof request and response (section 2.2.2.16), the number allowed is 0...1 (Choice of AppliesToInternal, AppliesToExternalKnown (section 2.2.3.12), and AppliesToExternalUnknown (section 2.2.3.13)).");

                                foreach (XmlNode appliesToInternal in appliesToInternalNodes)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R800");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R800
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToInternal.InnerXml),
                                        800,
                                        @"[In AppliesToInternal] The AppliesToInternal element is an empty tag element, meaning it has no value or data type.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R801");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R801
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToInternal.InnerXml),
                                        801,
                                        @"[In AppliesToInternal] It [AppliesToInternal element] is distinguished only by the presence or absence of the <AppliesToInternal/> tag.");
                                }
                            }
                            #endregion

                            #region Capture code for AppliesToExternalKnown
                            if (appliesToExternalKnownNodes.Count > 0)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1067");

                                // If the schema validation result is true and AppliesToExternalKnown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1067,
                                    @"[In AppliesToExternalKnown] Element AppliesToExternalKnown in Settings command Oof request and response (section 2.2.2.17), the parent element is OofMessage (section 2.2.3.117).");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1068");

                                // If the schema validation result is true and AppliesToExternalKnown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1068,
                                    @"[In AppliesToExternalKnown] None [Element  AppliesToExternalKnown in Settings command Oof request and response (section 2.2.2.17) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1069");

                                // If the schema validation result is true and AppliesToExternalKnown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1069,
                                    @"[In AppliesToExternalKnown] None [Element  AppliesToExternalKnown in Settings command Oof request and response (section 2.2.2.17), the data tye is None.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1070");

                                // If the schema validation result is true and AppliesToExternalKnown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1070,
                                    @"[In AppliesToExternalKnown] Element AppliesToExternalKnown in Settings command Oof request and response (section 2.2.2.17), the number allowed is 0...1 (Choice of AppliesToInternal (section 2.2.3.14), AppliesToExternalKnown, and AppliesToExternalUnknown (section 2.2.3.13)).");

                                foreach (XmlNode appliesToExternalKnown in appliesToExternalKnownNodes)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R781");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R781
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToExternalKnown.InnerXml),
                                        781,
                                        @"[In AppliesToExternalKnown] The AppliesToExternalKnown element is an empty tag element, meaning it [AppliesToExternalKnown element] has no value or data type.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R782");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R782
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToExternalKnown.InnerXml),
                                        782,
                                        @"[In AppliesToExternalKnown] It [AppliesToExternalKnown element] is distinguished only by the presence or absence of the <AppliesToExternalKnown/> tag.");
                                }
                            }
                            #endregion

                            #region Capture code for AppliesToExternalUnknown
                            if (appliesToExternalUnknownNodes.Count > 0)
                            {
                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1071");

                                // If the schema validation result is true and AppliesToExternalUnknown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1071,
                                    @"[In AppliesToExternalUnknown] Element AppliesToExternalUnknown in Settings command Oof request and response (section 2.2.2.17), the parent element is
OofMessage (section 2.2.3.117)");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1072");

                                // If the schema validation result is true and AppliesToExternalUnknown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1072,
                                    @"[In AppliesToExternalUnknown] None [Element AppliesToExternalUnknown in Settings command Oof request and response (section 2.2.2.17) has no child element.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1073");

                                // If the schema validation result is true and AppliesToExternalUnknown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1073,
                                    @"[In AppliesToExternalUnknown] None [Element AppliesToExternalUnknown in Settings command Oof request and response (section 2.2.2.17), the data type is None.]");

                                // Add the debug information.
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1074");

                                // If the schema validation result is true and AppliesToExternalUnknown exists, this requirement can be verified.
                                Site.CaptureRequirement(
                                    1074,
                                    @"[In AppliesToExternalUnknown] Element AppliesToExternalUnknown in Settings command Oof request and response (section 2.2.2.17), the number allowed is 0...1 (Choice of AppliesToInternal (section 2.2.3.14), AppliesToExternalKnown (section 2.2.3.12), and AppliesToExternalUnknown).");

                                foreach (XmlNode appliesToExternalUnknown in appliesToExternalUnknownNodes)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R791");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R791
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToExternalUnknown.InnerXml),
                                        791,
                                        @"[In AppliesToExternalUnknown] The AppliesToExternalUnknown element is an empty tag element, meaning it has no value or data type.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R792");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R792
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty(appliesToExternalUnknown.InnerXml),
                                        792,
                                        @"[In AppliesToExternalUnknown] It [AppliesToExternalUnknown element] is distinguished only by the presence or absence of the <AppliesToExternalUnknown/> tag.");
                                }
                            }
                            #endregion
                        }
                    }
                    #endregion
                }
                #endregion
            }
            #endregion

            #region Capture code for DeviceInformation
            if (settingsResponse.ResponseData.DeviceInformation != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1506");

                // If the schema validation result is true and DeviceInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1506,
                    @"[In DeviceInformation] Element DeviceInformation in Settings command response, the parent element is Settings.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1507");

                // If the schema validation result is true and DeviceInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1507,
                    @"[In DeviceInformation] Element DeviceInformation in Settings command response, the child element is  Status (section 2.2.3.167.14).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1508");

                // If the schema validation result is true and DeviceInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1508,
                    @"[In DeviceInformation] Element DeviceInformation in Settings command response, the data type is  container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1509");

                // If the schema validation result is true and DeviceInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1509,
                    @"[In DeviceInformation] Element DeviceInformation in Settings command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                Site.Assert.IsNotNull(settingsResponse.ResponseData.DeviceInformation.Status, "As child element of DeviceInformation, the Status element should not be null.");

                int status;

                Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.DeviceInformation.Status, out status), "As child element of DeviceInformation, the Status element should be an integer.");

                this.VerifyStatusElementForSettings();
            }
            #endregion

            #region Capture code for DevicePassword
            if (settingsResponse.ResponseData.DevicePassword != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1514");

                // If the schema validation result is true and DevicePassword is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1514,
                    @"[In DevicePassword] Element DevicePassword in Settings command response, the parent element is Settings.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1515");

                // If the schema validation result is true and DevicePassword is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1515,
                    @"[In DevicePassword] Element DevicePassword in Settings command response, the child element is Status (section 2.2.3.167.14).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1516");

                // If the schema validation result is true and DevicePassword is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1516,
                    @"[In DevicePassword] Element DevicePassword in Settings command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1517");

                // If the schema validation result is true and DevicePassword is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1517,
                    @"[In DevicePassword] Element DevicePassword in Settings command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                Site.Assert.IsNotNull(settingsResponse.ResponseData.DevicePassword.Status, "As a child element of DevicePassword, the Status element should not be null.");

                int status;

                Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.DevicePassword.Status, out status), "As a child element of DevicePassword, the Status element should be an integer.");

                this.VerifyStatusElementForSettings();

                Common.VerifyActualValues("Status(Settings)", AdapterHelper.ValidStatus(new string[] { "1", "2", "5", "7" }), settingsResponse.ResponseData.DevicePassword.Status, this.Site);

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4395");

                // Verify MS-ASCMD requirement: MS-ASCMD_R4395
                // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                Site.CaptureRequirement(
                    4395,
                    @"[In Status(Settings)] The following table lists the values [1,2,5,7] for Status in a Settings command DevicePassword Set response.");
            }
            #endregion

            #region Capture code for UserInformation
            if (settingsResponse.ResponseData.UserInformation != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2903");

                // If the schema validation result is true and UserInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2903,
                    @"[In UserInformation] Element UserInformation in Settings command response, the parent element is Settings.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2904");

                // If the schema validation result is true and UserInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2904,
                    @"[In UserInformation]  Element UserInformation in Settings command response, the child elements are Get, Status (section 2.2.3.167.14).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2905");

                // If the schema validation result is true and UserInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2905,
                    @"[In UserInformation]  Element UserInformation in Settings command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2906");

                // If the schema validation result is true and UserInformation is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2906,
                    @"[In UserInformation]  Element UserInformation in Settings command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();

                #region Capture code for Status
                Site.Assert.IsNotNull(settingsResponse.ResponseData.UserInformation.Status, "As child element of UserInformation, the Status element should not be null.");

                int status;

                Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.UserInformation.Status, out status), "As child element of UserInformation, the Status element should be an integer.");

                this.VerifyStatusElementForSettings();
                #endregion

                #region Capture code for Get
                Site.Assert.IsNotNull(settingsResponse.ResponseData.UserInformation.Get, "As a child element of UserInformation, the Get element should not be null.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1755");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1755,
                    @"[In Get] Element Get in Settings command UserInformation response, the parent element is UserInformation.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1758");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1758,
                    @"[In Get] Element Get in Settings command UserInformation response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1759");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1759,
                    @"[In Get] Element Get in Settings command UserInformation response, the number allowed is 1…1 (required).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5854");

                // If the schema validation result is true and Get is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5854,
                    @"[In Get] The Get element is a required child element of the UserInformation element in Settings command UserInformation requests and responses.");

                bool hasEmailAddresses = false;
                bool hasAccounts = false;
                #region Capture code for EmailAddresses
                if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") || Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                {
                    hasEmailAddresses = true;
                    if (settingsResponse.ResponseData.UserInformation.Get.EmailAddresses != null)
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1597");

                        // If the schema validation result is true and EmailAddresses is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1597,
                            @"[In EmailAddresses] Element EmailAddresses in Settings command UserInformation response (section 2.2.2.16), the data type is container ([MS-ASDTYPE] section 2.2).");

                        this.VerifyEmailAddresses(settingsResponse.ResponseData.UserInformation.Get.EmailAddresses);
                    }
                }
                #endregion

                #region Capture code for Accounts
                if (settingsResponse.ResponseData.UserInformation.Get.Accounts != null && settingsResponse.ResponseData.UserInformation.Get.Accounts.Length > 0)
                {
                    hasAccounts = true;

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1023");

                    // If the schema validation result is true and Accounts is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1023,
                        @"[In Accounts] Element Accounts in Settings command UserInformation response (section 2.2.2.17), the parent element is Get (section 2.2.3.79).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1025");

                    // If the schema validation result is true and Accounts is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1025,
                        @"[In Accounts] Element Accounts in Settings command UserInformation response (section 2.2.2.17), the data type is container ([MS-ASDTYPE] section 2.2).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1026");

                    // If the schema validation result is true and Accounts is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1026,
                        @"[In Accounts] Element Accounts in Settings command UserInformation response (section 2.2.2.17), the number allowed is 0…1 (optional).");

                    this.VerifyContainerDataType();

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R713");

                    // If the schema validation result is true and Account is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        713,
                        @"[In Account] The Account element is a required child element of the Accounts element in Settings command responses that contains all account information associated with a single account.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R998");

                    // If the schema validation result is true and Account is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        998,
                        @"[In Account] Element Account in Settings command UserInformation response (section 2.2.2.17), the parent element is Accounts (section 2.2.3.5).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1001");

                    // If the schema validation result is true and there is any Account element, this requirement can be verified.
                    Site.CaptureRequirement(
                        1001,
                        @"[In Account] Element Account in Settings command UserInformation response (section 2.2.2.17), the number allowed is 1…N (required).");

                    foreach (AccountsAccount account in settingsResponse.ResponseData.UserInformation.Get.Accounts)
                    {
                        Site.Assert.IsNotNull(account, "The Account element should not be null.");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1024");

                        // If the schema validation result is true and Accounts contains Account, this requirement can be verified.
                        Site.CaptureRequirement(
                            1024,
                            @"[In Accounts] Element Accounts in Settings command UserInformation response (section 2.2.2.17), the child element is Account (section 2.2.3.2).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R999");

                        // If the schema validation result is true and Account is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            999,
                            @"[In Account] The element Account in Settings command UserInformation response (section 2.2.2.17), the child elements are AccountId (section 2.2.3.3.2), AccountName (section 2.2.3.4), UserDisplayName (section 2.2.3.181), SendDisabled (section 2.2.3.151), EmailAddresses (section 2.2.3.54).");

                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1000");

                        // If the schema validation result is true and Account is not null, this requirement can be verified.
                        Site.CaptureRequirement(
                            1000,
                            @"[In Account] Element Account in Settings command UserInformation response (section 2.2.2.17), the data type is container ([MS-ASDTYPE] section 2.2).");

                        if (account.EmailAddresses != null)
                        {
                            this.VerifyEmailAddresses(account.EmailAddresses);
                        }

                        this.VerifyContainerDataType();
                    }
                }
                #endregion
                #endregion Capture code for Get
            }
            #endregion

            #region Capture code for Status
            Site.Assert.IsNotNull(settingsResponse.ResponseData.Status, "The Status element should not be null.");

            int statusForSettings;

            Site.Assert.IsTrue(int.TryParse(settingsResponse.ResponseData.Status, out statusForSettings), "The Status element should be an integer.");

            this.VerifyStatusElementForSettings();

            Common.VerifyActualValues("Status(Settings)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "7" }), settingsResponse.ResponseData.Status, this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4386");

            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                 4386,
                 @"[In Status(Settings)] The following table lists the valid values [1,2,3,4,5,6,7] for the Status element as the child element of the Settings element in the Settings command response.");
            #endregion
        }

        /// <summary>
        /// Verify EmailAddresses element
        /// </summary>
        /// <param name="emailAddresses">The serialized object of EmailAddresses element class</param>
        private void VerifyEmailAddresses(EmailAddresses emailAddresses)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1595");

            // If the schema validation result is true and EmailAddresses is not null, this requirement can be verified.
            Site.CaptureRequirement(
                 1595,
                 @"[In EmailAddresses] Element EmailAddresses in Settings command UserInformation response (section 2.2.2.16),the parent elements are Account<32> (section 2.2.3.2), Get<33> (section 2.2.3.75).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1598");

            // If the schema validation result is true and EmailAddresses is not null, this requirement can be verified.
            Site.CaptureRequirement(
                 1598,
                 @"[In EmailAddresses] Element EmailAddresses in Settings command UserInformation response (section 2.2.2.16), the number allowed is 0...1 (optional).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3962");

            Site.CaptureRequirementIfIsNotNull(
                            emailAddresses.SMTPAddress,
                            3962,
                            @"[In SMTPAddress] The SMTPAddress element is a required child element of the EmailAddresses element in Settings command responses that specifies one of the user's email addresses.");

            bool hasPrimarySmtpAddress = false;
            if (emailAddresses.PrimarySmtpAddress != null)
            {
                hasPrimarySmtpAddress = true;

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2392");

                // If the schema validation result is true and PrimarySmtpAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                     2392,
                     @"[In PrimarySmtpAddress] Element PrimarySmtpAddress in Settings command UserInformation response (section 2.2.2.17), the parent element is EmailAddresses (section 2.2.3.54).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2393");

                // If the schema validation result is true and PrimarySmtpAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                     2393,
                     @"[In PrimarySmtpAddress] None [Element PrimarySmtpAddress in Settings command UserInformation response (section 2.2.2.17) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2394");

                // If the schema validation result is true and PrimarySmtpAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                     2394,
                     @"[In PrimarySmtpAddress] Element PrimarySmtpAddress in Settings command UserInformation response (section 2.2.2.17), the data type is string ([MS-ASDTYPE] section 2.7).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2395");

                // If the schema validation result is true and PrimarySmtpAddress is not null, this requirement can be verified.
                Site.CaptureRequirement(
                     2395,
                     @"[In PrimarySmtpAddress] Element PrimarySmtpAddress in Settings command UserInformation response (section 2.2.2.17), the number allowed is 0…1 (optional).");
            }

            if (hasPrimarySmtpAddress)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1596");

                // If the schema validation result is true and R3962 is captured, when PrimarySmtpAddress not null, this requirement can be verified.
                Site.CaptureRequirement(
                     1596,
                     @"[In EmailAddresses] Element EmailAddresses in Settings command UserInformation response (section 2.2.2.16), the child elements are SMTPAddress (section 2.2.3.156), PrimarySmtpAddress<34> (section 2.2.3.127).");
            }
        }
        #endregion

        #region Capture code for SmartForward command
        /// <summary>
        /// This method is used to verify the SmartForward response related requirements.
        /// </summary>
        /// <param name="smartForwardResponse">SmartForward command response.</param>
        private void VerifySmartForwardCommand(SmartForwardResponse smartForwardResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(smartForwardResponse.ResponseData, "The SmartForward element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3957");

            // If the schema validation result is true and SmartForward is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3957,
                @"[In SmartForward] The SmartForward element is a required element in SmartForward command requests and responses that identifies the body of the HTTP POST as containing a SmartForward command (section 2.2.2.18).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2639");

            // If the schema validation result is true and SmartForward is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2639,
                @"[In SmartForward]None [Element SmartForward in SmartForward command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2640");

            // If the schema validation result is true and SmartForward is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2640,
                @"[In SmartForward] Element SmartForward in SmartForward command response, the child eleemnt is Status (section 2.2.3.167.15).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2641");

            // If the schema validation result is true and SmartForward is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2641,
                @"[In SmartForward] Element SmartForward in SmartForward command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2642");

            // If the schema validation result is true and SmartForward is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2642,
                @"[In SmartForward] Element SmartForward in SmartForward command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();

            if (smartForwardResponse.ResponseData.Status != null)
            {
                int status;

                Site.Assert.IsTrue(int.TryParse(smartForwardResponse.ResponseData.Status, out status), "The status element in SmartForward response should be an integer.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2759");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2759,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartForward command response, the parent  element is SmartForward (section 2.2.3.159).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2760");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2760,
                    @"[In Status(SmartForward and SmartReply)] None [Element Status in SmartForward command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2761");

                // If the schema validation result is true, Status(SmartForward and SmartReply) is an integer, this requirement can be verified.
                Site.CaptureRequirement(
                    2761,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartForward command response, the data type is integer ([MS-ASDTYPE] section 2.6).");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2762");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2762,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartForward command response, the number allowed is 0...1 (optional).");
            }
        }
        #endregion

        #region Capture code for SmartReply command
        /// <summary>
        /// This method is used to verify the SmartReply response related requirements.
        /// </summary>
        /// <param name="smartReplyResponse">SmartReply command response.</param>
        private void VerifySmartReplyCommand(SmartReplyResponse smartReplyResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(smartReplyResponse.ResponseData, "The SmartReply element should not be null.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3960");

            // If the schema validation result is true and SmartReply is not null, this requirement can be verified.
            Site.CaptureRequirement(
                3960,
                @"[In SmartReply] The SmartReply element is a required element in SmartReply command requests and responses that identifies the body of the HTTP POST as containing a SmartReply command (section 2.2.2.18). ");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2647");

            // If the schema validation result is true and SmartReply is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2647,
                @"[In SmartReply] None [Element SmartReply in SmartReply command reply has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2648");

            // If the schema validation result is true and SmartReply is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2648,
                @"[In SmartReply] Element SmartReply in SmartReply command reply, the child element is Status (section 2.2.3.167.15).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2649");

            // If the schema validation result is true and SmartReply is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2649,
                @"[In SmartReply] Element SmartReply in SmartReply command reply, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2650");

            // If the schema validation result is true and SmartReply is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2650,
                @"[In SmartReply] Element SmartReply in SmartReply command reply, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();

            if (smartReplyResponse.ResponseData.Status != null)
            {
                int status;

                Site.Assert.IsTrue(int.TryParse(smartReplyResponse.ResponseData.Status, out status), "The status element in SmartReply response should be an integer.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2763");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2763,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartReply command response, the parent element is SmartReply (section 2.2.3.160).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2764");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2764,
                    @"[In Status(SmartForward and SmartReply)] None [Element Status in SmartReply command response has no child element .]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2765");

                // If the schema validation result is true, Status(SmartForward and SmartReply) is an integer, this requirement can be verified.
                Site.CaptureRequirement(
                    2765,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartReply command response, the data type is integer.");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2766");

                // If the schema validation result is true and Status is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    2766,
                    @"[In Status(SmartForward and SmartReply)] Element Status in SmartReply command response, the number allowed is 0...1 (optional).");
            }
        }
        #endregion

        #region Capture code for Sync command
        /// <summary>
        /// This method is used to verify the Sync response related requirements.
        /// </summary>
        /// <param name="syncResponse">Sync command response.</param>
        private void VerifySyncCommand(SyncResponse syncResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(syncResponse.ResponseData, "The Sync element should not be null.");

            #region Capture code for Sync
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4558");

            // If the schema validation result is true and Sync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                4558,
                @"[In Sync] The Sync element is a required element in Sync command requests and responses that identifies the body of the HTTP POST as containing a Sync command (section 2.2.2.20).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2795");

            // If the schema validation result is true and Sync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2795,
                @"[In Sync] None [Element Sync in Sync command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2796");

            // If the schema validation result is true and Sync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2796,
                @"[In Sync] Element Sync in Sync command response, the child elements are Collections (section 2.2.3.31.2), Limit (section 2.2.3.92), Status (section 2.2.3.167.16).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2797");

            // If the schema validation result is true and Sync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2797,
                @"[In Sync] Element Sync in Sync command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2798");

            // If the schema validation result is true and Sync is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2798,
                @"[In Sync] Element Sync in Sync command response, the number allowed is 1...1 (required).");

            this.VerifyContainerDataType();
            #endregion

            #region Capture code for Status
            if (!string.IsNullOrEmpty(syncResponse.ResponseData.Status))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5585");

                // If the schema validation result is true and Status(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5585,
                    @"[In Status(Sync)] Element Status in Sync command response, the parent element is Sync (section 2.2.3.170).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5586");

                // If the schema validation result is true and Status(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5586,
                    @"[In Status(Sync)] None [Element Status in Sync command response has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5587");

                // If the schema validation result is true and Status(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5587,
                    @"[In Status(Sync)] Element Status in Sync command response, the data type is unsignedByte.");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5588");

                // If the schema validation result is true and Status(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    5588,
                    @"[In Status(Sync)] Element Status in Sync command response, the number allowed is 0…1 (optional).");
            }
            #endregion

            if (!string.IsNullOrEmpty(syncResponse.ResponseData.Status) && (syncResponse.ResponseData.Status == "14" || syncResponse.ResponseData.Status == "15"))
            {
                Site.Assert.IsNotNull((string)syncResponse.ResponseData.Item, "The Limit element should not be null when the Status is 14 or 15.");

                #region Capture code for Limit
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1828");

                // If the schema validation result is true and Limit is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1828,
                    @"[In Limit] Element Limit in Sync command response (section 2.2.2.20), the parent element is Sync (section 2.2.3.170).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1829");

                // If the schema validation result is true and Limit is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1829,
                    @"[In Limit] None [Element Limit in Sync command response (section 2.2.2.20) has no child element.]");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1830");

                // If the schema validation result is true and Limit is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1830,
                    @"[In Limit] Element Limit in Sync command response (section 2.2.2.20), the data type is integer ([MS-ASDTYPE] section 2.6).");

                this.VerifyIntegerDataType();

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1831");

                // If the schema validation result is true and Limit is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1831,
                    @"[In Limit] Element Limit in Sync command response (section 2.2.2.20), the number allowed is 0...1 (optional).");
                #endregion
            }
            else if ((SyncCollections)syncResponse.ResponseData.Item != null)
            {
                SyncCollections syncCollections = (SyncCollections)syncResponse.ResponseData.Item;

                #region Capture code for Collections
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1415");

                // If the schema validation result is true and Collections(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1415,
                    @"[In Collections(Sync)] Element Commands in Sync command response, the parent element is Sync.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1416");

                // If the schema validation result is true and Collections(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1416,
                    @"[In Collections(Sync)] Element Commands in Sync command response, the child element is Collection.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1417");

                // If the schema validation result is true and Collections(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1417,
                    @"[In Collections(Sync)] Element Commands in Sync command response, the data type is container.");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1418");

                // If the schema validation result is true and Collections(Sync) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1418,
                    @"[In Collections(Sync)] Element Commands in Sync command response, the number allowed is 0...1 (optional).");

                this.VerifyContainerDataType();
                #endregion

                #region Capture code for Collection
                if (syncCollections.Collection != null && syncCollections.Collection.Length > 0)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1363");

                    // If the schema validation result is true and Collection(Sync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1363,
                        @"[In Collection(Sync)] Element Collection in Sync command response, the parent element is Collections.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1364");

                    // If the schema validation result is true and Collection(Sync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1364,
                        @"[In Collection(Sync)] Element Collection in Sync command response, the child elements are Class, SyncKey, CollectionId, Status (section 2.2.3.167.16), MoreAvailable (section 2.2.3.110), Commands, Responses (section 2.2.3.145).");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1365");

                    // If the schema validation result is true and Collection(Sync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1365,
                        @"[In Collection(Sync)] Element Collection in Sync command response, the data type is container.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1366");

                    // If the schema validation result is true and Collection(Sync) is not null, this requirement can be verified.
                    Site.CaptureRequirement(
                        1366,
                        @"[In Collection(Sync)] Element Collection in Sync command response, the number allowed is 0…N (optional).");

                    this.VerifyContainerDataType();

                    foreach (SyncCollectionsCollection collection in syncCollections.Collection)
                    {
                        if (collection.ItemsElementName != null && collection.ItemsElementName.Length > 0)
                        {
                            bool hasStatus = false;
                            bool hasSyncKey = false;
                            bool hasCollectionId = false;

                            for (int j = 0; j < collection.ItemsElementName.Length; j++)
                            {
                                #region Capture code for Status
                                if (collection.ItemsElementName[j] == ItemsChoiceType10.Status && collection.Items[j] != null)
                                {
                                    hasStatus = true;

                                    Common.VerifyActualValues("Status(Sync)", AdapterHelper.ValidStatus(new string[] { "1", "3", "4", "5", "6", "7", "8", "9", "12", "13", "14", "15", "16" }), collection.Items[j].ToString(), this.Site);

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4420");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R4420
                                    // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        4420,
                                        @"[In Status(Sync)] The following table lists the status codes [1,3,4,5,6,7,8,9,12,13,14,15,16] for the Sync command (section 2.2.2.20). For information about the scope of the status value and for status values common to all ActiveSync commands, see section 2.2.4.");

                                    this.VerifyStatusElementForSync();
                                }
                                #endregion

                                #region Capture code for SyncKey
                                if (collection.ItemsElementName[j] == ItemsChoiceType10.SyncKey && collection.Items[j] != null)
                                {
                                    hasSyncKey = true;

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2839");

                                    // If the schema validation result is true and SyncKey(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2839,
                                        @"[In SyncKey(Sync)] Element SyncKey in Sync command response, the parent element is Collection (section 2.2.3.29.2),");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2840");

                                    // If the schema validation result is true and SyncKey(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2840,
                                        @"[In SyncKey(Sync)] None [Element SyncKey in Sync command response has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2841");

                                    // If the schema validation result is true and SyncKey(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        2841,
                                        @"[In SyncKey(Sync)] Element SyncKey in Sync command response, the data type is string.");

                                    this.VerifyStringDataType();
                                }
                                #endregion

                                #region Capture code for CollectionId
                                if (collection.ItemsElementName[j] == ItemsChoiceType10.CollectionId && collection.Items[j] != null)
                                {
                                    hasCollectionId = true;

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1400");

                                    // If the schema validation result is true and CollectionId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1400,
                                        @"[In CollectionId(Sync)] Element CollectionId in Sync command response, the parent element is Collection (section 2.2.3.29.2).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1401");

                                    // If the schema validation result is true and CollectionId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1401,
                                        @"[In CollectionId(Sync)] None [Element CollectionId in Sync command response has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1402");

                                    // If the schema validation result is true and CollectionId(Sync) is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1402,
                                        @"[In CollectionId(Sync)] Element CollectionId in Sync command response, the data type is string.");

                                    this.VerifyStringDataType();

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5862, the length of CollectionId element is {0}", ((string)collection.Items[j]).Length);

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R5862
                                    Site.CaptureRequirementIfIsTrue(
                                        ((string)collection.Items[j]).Length <= 64,
                                        5862,
                                        @"[In CollectionId(Sync)] The CollectionId element value is not larger than 64 characters in length.");
                                }
                                #endregion

                                #region Capture code for MoreAvailable
                                if (collection.ItemsElementName[j] == ItemsChoiceType10.MoreAvailable)
                                {
                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1951");

                                    // If the schema validation result is true and MoreAvailable is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1951,
                                        @"[In MoreAvailable] Element MoreAvailable in Sync command response (section 2.2.2.20), the parent element is Collection (section 2.2.3.29.2).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1952");

                                    // If the schema validation result is true and MoreAvailable is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1952,
                                        @"[In MoreAvailable] None [Element MoreAvailable in Sync command response (section 2.2.2.20) has no child element.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1953");

                                    // If the schema validation result is true and MoreAvailable is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1953,
                                        @"[In MoreAvailable] None [Element MoreAvailable in Sync command response (section 2.2.2.20) has no data type.]");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1954");

                                    // If the schema validation result is true and MoreAvailable is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1954,
                                        @"[In MoreAvailable] Element MoreAvailable in Sync command response (section 2.2.2.20), the number allowed is 0...1 (optional).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3448");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R3448
                                    Site.CaptureRequirementIfIsTrue(
                                        string.IsNullOrEmpty((string)collection.Items[j]),
                                        3448,
                                        @"[In MoreAvailable] The MoreAvailable element is an empty tag element, meaning it has no value or data type.");
                                }
                                #endregion

                                #region Capture code for Commands
                                if (collection.ItemsElementName[j] == ItemsChoiceType10.Commands && collection.Items[j] != null)
                                {
                                    SyncCollectionsCollectionCommands commands = (SyncCollectionsCollectionCommands)collection.Items[j];

                                    bool hasAddCommand = commands.Add != null && commands.Add.Length > 0;
                                    bool hasDeleteCommand = commands.Delete != null && commands.Delete.Length > 0;
                                    bool hasChangeCommand = commands.Change != null && commands.Change.Length > 0;
                                    bool hasSoftDeleteCommand = commands.SoftDelete != null && commands.SoftDelete.Length > 0;

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2053");

                                    // Verify MS-ASCMD requirement: MS-ASCMD_R2053
                                    Site.CaptureRequirementIfIsTrue(
                                        hasAddCommand || hasDeleteCommand || hasChangeCommand || hasSoftDeleteCommand,
                                        2053,
                                        @"[In Commands] If it [Commands] is present, it MUST include at least one operation.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1423");

                                    // If the schema validation result is true and Commands is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1423,
                                        @"[In Commands] Element Commands in Sync command response, the parent element is Collection.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1425");

                                    // If the schema validation result is true and Commands is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1425,
                                        @"[In Commands] Element Commands in Sync command response, the child elements are Add, Delete, Change, SoftDelete (section 2.2.3.162).");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1426");

                                    // If the schema validation result is true and Commands is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1426,
                                        @"[In Commands] Element Commands in Sync command response, the data type is  container.");

                                    // Add the debug information.
                                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1427");

                                    // If the schema validation result is true and Commands is not null, this requirement can be verified.
                                    Site.CaptureRequirement(
                                        1427,
                                        @"[In Commands] Element Commands in Sync command response, the number allowed is 0...1 (optional).");

                                    this.VerifyContainerDataType();

                                    #region Capture code for SoftDelete
                                    if (hasSoftDeleteCommand)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2655");

                                        // If the schema validation result is true and SoftDelete is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2655,
                                            @"[In SoftDelete] Element SoftDelete in Sync command response (section 2.2.2.20),the parent element is Commands (section 2.2.3.32).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2656");

                                        // If the schema validation result is true and SoftDelete is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2656,
                                            @"[In SoftDelete] Element SoftDelete in Sync command response (section 2.2.2.20), the child element is ServerId (section 2.2.3.156.6).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2657");

                                        // If the schema validation result is true and SoftDelete is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2657,
                                            @"[In SoftDelete] Element SoftDelete in Sync command response (section 2.2.2.20), the data type is container ([MS-ASDTYPE] section 2.2).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2658");

                                        // If the schema validation result is true and SoftDelete is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            2658,
                                            @"[In SoftDelete] Element SoftDelete in Sync command response (section 2.2.2.20), the number allowed is 0...N (optional).");

                                        this.VerifyContainerDataType();

                                        foreach (SyncCollectionsCollectionCommandsSoftDelete softDelete in commands.SoftDelete)
                                        {
                                            Site.Assert.IsNotNull(softDelete.ServerId, "The ServerId element in Sync command response should not be null.");

                                            this.VerifyServerIdElementForSync(softDelete.ServerId);
                                        }
                                    }
                                    #endregion

                                    #region Capture code for Delete
                                    if (hasDeleteCommand)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1490");

                                        // If the schema validation result is true and Delete(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1490,
                                            @"[In Delete(Sync)] Element Delete in Sync command response, the parent element is Commands.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1491");

                                        // If the schema validation result is true and Delete(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1491,
                                            @"[In Delete(Sync)] Element Delete in Sync command response, the child element is ServerId, Class (section 2.2.3.27.5).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1492");

                                        // If the schema validation result is true and Delete(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1492,
                                            @"[In Delete(Sync)] Element Delete in Sync command response, the data type is  container.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1493");

                                        // If the schema validation result is true and Delete(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1493,
                                            @"[In Delete(Sync)] Element Delete in Sync command response, the number allowed is 0...N (optional).");

                                        this.VerifyContainerDataType();

                                        foreach (SyncCollectionsCollectionCommandsDelete delete in commands.Delete)
                                        {
                                            Site.Assert.IsNotNull(delete.ServerId, "The ServerId element in Sync command response should not be null.");

                                            this.VerifyServerIdElementForSync(delete.ServerId);
                                        }
                                    }
                                    #endregion

                                    #region Capture code for Add
                                    if (hasAddCommand)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1039");

                                        // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1039,
                                            @"[In Add(Sync)] Element Add in Sync command response, the parent element is Commands.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1040");

                                        // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1040,
                                            @"[In Add(Sync)] Element Add in Sync command response, the child elements are ServerId (section 2.2.3.156.7), ApplicationData (section 2.2.3.11).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1041");

                                        // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1041,
                                            @"[In Add(Sync)] Element Add in Sync command response, the data type is container.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1042");

                                        // If the schema validation result is true and Add(Sync) is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1042,
                                            @"[In Add(Sync)] Element Add in Sync command response, the number allowed is 0...N (optional).");

                                        this.VerifyContainerDataType();

                                        foreach (SyncCollectionsCollectionCommandsAdd add in commands.Add)
                                        {
                                            Site.Assert.IsNotNull(add.ServerId, "The ServerId element in Sync command response should not be null.");

                                            this.VerifyServerIdElementForSync(add.ServerId);

                                            Site.Assert.IsNotNull(add.ApplicationData, "The ApplicationData element in Sync command should not be null.");

                                            this.VerifyApplicationDataForSyncAddChange();
                                            for (int i = 0; i < add.ApplicationData.ItemsElementName.Length; i++)
                                            {
                                                if (add.ApplicationData.ItemsElementName[i].ToString() == "MeetingRequst")
                                                {
                                                    if (add.ApplicationData.Items[i] != null)
                                                    {
                                                        if (((MeetingRequest)add.ApplicationData.Items[i]).Forwardees != null)
                                                        {
                                                            VerifyForwardeesElementForSyncResponses();
                                                            foreach (ForwardeesForwardee forwardee in ((MeetingRequest)add.ApplicationData.Items[i]).Forwardees)
                                                            {
                                                                if (forwardee != null)
                                                                {
                                                                    VerifyForwardeeElementForSyncResponses();
                                                                    if (forwardee.Name != null)
                                                                    {
                                                                        VerifyForwardeeEmailElementForSyncResponses();
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (((MeetingRequest)add.ApplicationData.Items[i]).ProposedEndTime != null)
                                                        {
                                                            VerifyMeetingRequestProposedEndTimeElementForSyncResponses();
                                                        }
                                                        if (((MeetingRequest)add.ApplicationData.Items[i]).ProposedStartTime != null)
                                                        {
                                                            VerifyMeetingRequestProposedStartTimeElementForSyncResponses();
                                                        }
                                                    }
                                                }
                                                if (add.ApplicationData.ItemsElementName[i].ToString() == "Attendees")
                                                {
                                                    if (add.ApplicationData.Items[i] != null)
                                                    {
                                                        if (((Attendees)add.ApplicationData.Items[i]).Attendee != null)
                                                        {
                                                            VerifyForwardeesElementForSyncResponses();
                                                            foreach (AttendeesAttendee attende in ((Attendees)add.ApplicationData.Items[i]).Attendee)
                                                            {
                                                                if (attende.ProposedEndTime != null)
                                                                {
                                                                    VerifyAttendeeProposedEndTimeElementForSyncResponses();
                                                                }
                                                                if (attende.ProposedStartTime!=null)
                                                                {
                                                                    VerifyAttendeeProposedStartTimeElementForSyncResponses();
                                                                }
                                                            }
                                                        }                                                       
                                                    }

                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Capture code for Change
                                    if (hasChangeCommand)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1147");

                                        // If the schema validation result is true and Change is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1147,
                                            @"[In Change] Element Change in Sync command response, the parent element is Commands.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1148");

                                        // If the schema validation result is true and Change is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1148,
                                            @"[In Change] Element Change in Sync command response, the child elements are ServerId, ApplicationData, Class (section 2.2.3.27.5).");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1149");

                                        // If the schema validation result is true and Change is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1149,
                                            @"[In Change] Element Change in Sync command response, the data type is container.");

                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1150");

                                        // If the schema validation result is true and Change is not null, this requirement can be verified.
                                        Site.CaptureRequirement(
                                            1150,
                                            @"[In Change] Element Change in Sync command response, the number allowed is 0...N (optional).");

                                        this.VerifyContainerDataType();

                                        foreach (SyncCollectionsCollectionCommandsChange change in commands.Change)
                                        {
                                            Site.Assert.IsNotNull(change.ServerId, "The ServerId element in Sync command response should not be null.");

                                            this.VerifyServerIdElementForSync(change.ServerId);

                                            Site.Assert.IsNotNull(change.ApplicationData, "The ApplicationData element in Sync command should not be null.");

                                            this.VerifyApplicationDataForSyncAddChange();
                                            for (int i = 0; i < change.ApplicationData.ItemsElementName.Length; i++)
                                            {
                                                if (change.ApplicationData.ItemsElementName[i].ToString() == "MeetingRequst")
                                                {
                                                    if (change.ApplicationData.Items[i] != null)
                                                    {
                                                        if (((MeetingRequest)change.ApplicationData.Items[i]).Forwardees != null)
                                                        {
                                                            VerifyForwardeesElementForSyncResponses();

                                                            foreach (ForwardeesForwardee forwardee in ((MeetingRequest)change.ApplicationData.Items[i]).Forwardees)
                                                            {
                                                                if (forwardee!=null)
                                                                {
                                                                    VerifyForwardeeElementForSyncResponses();
                                                                }                                                                
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                }
                                #endregion

                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.LoadXml(syncResponse.ResponseDataXML);

                                if (xmlDoc.DocumentElement.HasChildNodes)
                                {
                                    #region Capture code for Change
                                    XmlNodeList changeNodes = xmlDoc.DocumentElement.GetElementsByTagName("Change");

                                    if (changeNodes != null && changeNodes.Count > 0)
                                    {
                                        // Add the debug information.
                                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R3836");

                                        // Verify MS-ASCMD requirement: MS-ASCMD_R3836
                                        Site.CaptureRequirement(
                                            3836,
                                            @"[In Responses] If [ Responses element is ] present, it MUST include at least one child element.");
                                    }
                                    #endregion

                                    #region Capture code for Responses/Response
                                    XmlNodeList responsesNodes = xmlDoc.DocumentElement.GetElementsByTagName("Responses");
                                    XmlNodeList responseNodes = xmlDoc.DocumentElement.GetElementsByTagName("Response");

                                    if (responsesNodes != null && responsesNodes.Count > 0)
                                    {
                                        foreach (XmlNode responsesNode in responsesNodes)
                                        {
                                            if (responsesNode.ParentNode.Name.Equals("Collection", StringComparison.CurrentCultureIgnoreCase) && responsesNode.ParentNode.ParentNode != null && responsesNode.ParentNode.ParentNode.Name.Equals("Collections", StringComparison.CurrentCultureIgnoreCase) && responsesNode.ParentNode.ParentNode.ParentNode != null && responsesNode.ParentNode.ParentNode.ParentNode.Name.Equals("Sync", StringComparison.CurrentCultureIgnoreCase))
                                            {
                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2504");

                                                // If the schema validation result is true and Responses is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    2504,
                                                    @"[In Responses] Element Responses in Sync command response (section 2.2.2.20), the parent element is Collection (section 2.2.3.29.2).");

                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2506");

                                                // If the schema validation result is true and Responses is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    2506,
                                                    @"[In Responses] Element Responses in Sync command response (section 2.2.2.19), the data  type is container ([MS-ASDTYPE] section 2.2).");

                                                // Add the debug information.
                                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2507");

                                                // If the schema validation result is true and Responses is not null, this requirement can be verified.
                                                Site.CaptureRequirement(
                                                    2507,
                                                    @"[In Responses] Element Responses in Sync command response (section 2.2.2.19), the number allowed is 0...1 (optional).");

                                                this.VerifyContainerDataType();

                                                this.VerifyElementsForResponses(responsesNode as XmlElement);
                                            }
                                        }
                                    }

                                    if (responseNodes != null && responseNodes.Count > 0)
                                    {
                                        foreach (XmlNode responseNode in responseNodes)
                                        {
                                            if (responseNode.ParentNode.Name.Equals("Collection", StringComparison.CurrentCultureIgnoreCase) && responseNode.ParentNode.ParentNode != null && responseNode.ParentNode.ParentNode.Name.Equals("Collections", StringComparison.CurrentCultureIgnoreCase) && responseNode.ParentNode.ParentNode.ParentNode != null && responseNode.ParentNode.ParentNode.ParentNode.Name.Equals("Sync", StringComparison.CurrentCultureIgnoreCase))
                                            {
                                                this.VerifyElementsForResponses(responseNode as XmlElement);
                                            }
                                        }
                                    }
                                    #endregion
                                }
                            }

                            Site.Assert.IsTrue(hasStatus, "The Status in Collection should not be null.");
                            Site.Assert.IsTrue(hasSyncKey, "The SyncKey in Collection should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4596");

                            // If the schema validation result is true and SyncKey(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                4596,
                                @"[In SyncKey(Sync)] The SyncKey element is a required child element of the Collection element in Sync command requests and responses that contains a value that is used by the server to mark the synchronization state of a collection.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2842");

                            // If the schema validation result is true and SyncKey(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2842,
                                @"[In SyncKey(Sync)] Element SyncKey in Sync command response, the number allowed is 1…1 (required).");

                            Site.Assert.IsTrue(hasCollectionId, "The CollectionId in Collection should not be null.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2033");

                            // If the schema validation result is true and CollectionId(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                2033,
                                @"[In CollectionId(Sync)] The CollectionId element is a required child element of the Collection element in Sync command requests and responses that specifies the server ID of the folder to be synchronized.");

                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1403");

                            // If the schema validation result is true and CollectionId(Sync) is not null, this requirement can be verified.
                            Site.CaptureRequirement(
                                1403,
                                @"[In CollectionId(Sync)] Element CollectionId in Sync command response, the number allowed is 1…1 (required).");
                        }
                    }
                }
                #endregion
            }
        }

        #endregion

        #region Capture code for ValidateCert command
        /// <summary>
        /// This method is used to verify the ValidateCert response related requirements.
        /// </summary>
        /// <param name="validateCertResponse">ValidateCert command response.</param>
        private void VerifyValidateCertCommand(ValidateCertResponse validateCertResponse)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(validateCertResponse.ResponseData, "The ValidateCert element should not be null.");
            byte validateCertStatusValue = AdapterHelper.PickUpRootStatusValueFromXMLString(validateCertResponse.ResponseDataXML);
            string certificateStatusValue = AdapterHelper.GetValidateCertStatusCode(validateCertResponse);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R4727");

            // If the schema validation result is true and ValidateCert is not null, this requirement can be verified.
            Site.CaptureRequirement(
                4727,
                @"[In ValidateCert] The ValidateCert element is a required element in ValidateCert command requests and responses that identifies the body of the HTTP POST as containing a ValidateCert command (section 2.2.2.21).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2923");

            // If the schema validation result is true and ValidateCert is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2923,
                @"[In ValidateCert] None [Element ValidateCert in ValidateCert command response has no parent element.]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2924");

            // If the schema validation result is true and ValidateCert is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2924,
                @"[In ValidateCert] Element ValidateCert in ValidateCert command response, the child elements are Status (section 2.2.3.167.17), Certificate (section 2.2.3.19.2).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2925");

            // If the schema validation result is true and ValidateCert is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2925,
                @"[In ValidateCert] Element ValidateCert in ValidateCert command response, the data type is container.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R2926");

            // If the schema validation result is true and ValidateCert is not null, this requirement can be verified.
            Site.CaptureRequirement(
                2926,
                @"[In ValidateCert] Element ValidateCert in ValidateCert command response, the number allowed is 1…1 (required).");

            this.VerifyContainerDataType();

            Common.VerifyActualValues("Status(ValidateCert)", AdapterHelper.ValidStatus(new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17" }), validateCertStatusValue.ToString(), this.Site);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R5389");

            // Verify MS-ASCMD requirement: MS-ASCMD_R5389
            // If above Common.VerifyActualValues method is not failed, this requirement can be verified.
            Site.CaptureRequirement(
                5389,
                @"[In Status(ValidateCert)] The following table lists the status codes [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17] that apply to certificate validation for the ValidateCert command (section 2.2.2.21).");

            this.VerifyStatusElementForValidateCert();

            if (certificateStatusValue != null)
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1119");

                // If the schema validation result is true and Certificate(ValidateCert) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1119,
                    @"[In Certificate(ValidateCert)] Element Certificate in ValidateCert command response, the parent element is ValidateCert (section 2.2.3.185).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1120");

                // If the schema validation result is true and Certificate(ValidateCert) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1120,
                    @"[In Certificate(ValidateCert)] Element Certificate in ValidateCert command response, the child element is Status (section 2.2.3.167.17).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1121");

                // If the schema validation result is true and Certificate(ValidateCert) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1121,
                    @"[In Certificate(ValidateCert)] Element Certificate in ValidateCert command response, the data type is container ([MS-ASDTYPE] section 2.2).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASCMD_R1122");

                // If the schema validation result is true and Certificate(ValidateCert) is not null, this requirement can be verified.
                Site.CaptureRequirement(
                    1122,
                    @"[In Certificate(ValidateCert)] Element Certificate in ValidateCert command response, the number allowed is 0...N (optional).");

                this.VerifyContainerDataType();
                this.VerifyStatusElementForValidateCert();
            }
        }
        #endregion

        #region Capture code for GetHierarchy command
        /// <summary>
        /// This method is used to verify the GetHierarchy response related requirements.
        /// </summary>
        /// <param name="response">GetHierarchy command response.</param>
        private void VerifyGetHierarchyCommand(GetHierarchyResponse response)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(response.ResponseData, "The Folders element should not be null.");
            
            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6046,
                @"[In GetHierarchy] That is, there is no such top-level element called ""GetHierarchy"" that identifies the body of the HTTP POST as containing a GetHierarchy command. ");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6047,
                @"[In GetHierarchy] Instead, the Folders element is the top-level element.");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6600,
                @"[In Folders(GetHierarchy)] None [Element Folders in GetHierarchy command response (section 2.2.2.7) has no parent element.]");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6601,
                @"[In Folders(GetHierarchy)] Element Folders in GetHierarchy command response (section 2.2.2.7), the child element is Folder (section 2.2.3.66).");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6602,
                @"[In Folders(GetHierarchy)] Element Folders in GetHierarchy command response (section 2.2.2.7), the data type is container ([MS-ASDTYPE] section 2.2).");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6603,
                @"[In Folders(GetHierarchy)] Element Folders in GetHierarchy command response (section 2.2.2.7), the number allowed is 1...1 (required).");

            foreach (FoldersFolder folder in response.ResponseData.Folder)
            {
                this.VerifyFolderElement(folder);
            }
        }

        /// <summary>
        /// Verify the requirements about folder element.
        /// </summary>
        /// <param name="folder">The folder element.</param>
        private void VerifyFolderElement(FoldersFolder folder)
        {
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "The schema validation result should be true.");
            Site.Assert.IsNotNull(folder, "The folder element should not be null.");

            // If the schema validation result is true and response is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6567,
                @"[In Folder(GetHierarchy)] The Folder element is a required child element of the Folders element in GetHierarchy command responses that contains details about a folder.");

            // If the schema validation result is true and folder is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6569,
                @"[In folder(GetHierarchy)] Element Folder in GetHierarchy command response (section 2.2.2.7), the parent element is Folders (section 2.2.3.70.1).");

            // If the schema validation result is true and folder is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6570,
                @"[In folder(GetHierarchy)] Element Folder in GetHierarchy command response (section 2.2.2.7), the child elements are DisplayName (section 2.2.3.47.4), ServerId (section 2.2.3.156.5), Type (section 2.2.3.176.4), ParentId (section 2.2.3.123.4).");

            // If the schema validation result is true and folder is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6571,
                @"[In Folder(GetHierarchy)] Element Folder in GetHierarchy command response (section 2.2.2.7), the data type is container ([MS-ASDTYPE] section 2.2).");

            // If the schema validation result is true and folder is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6572,
                @"[In Folder(GetHierarchy)] Element Folder in GetHierarchy command response (section 2.2.2.7), the number allowed is 1...N (required).");

            Site.Assert.IsTrue(!string.IsNullOrEmpty(folder.DisplayName), "The DisplayName element should not be null.");

            // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6452,
                @"[In DisplayName(GetHierarchy)] The DisplayName element is a required child element of the Folder element in GetHierarchy command responses that specifies the display name of the folder.");

            // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6456,
                @"[In DisplayName(GetHierarchy)] Element DisplayName in GetHierarchy command response (section 2.2.2.7), the data type is string ([MS-ASDTYPE] section 2.7).");

            // If the schema validation result is true and DisplayName is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6457,
                @"[In DisplayName(GetHierarchy)] Element DisplayName in GetHierarchy command response (section 2.2.2.7), the number allowed is 1...1 (required).");

            Site.Assert.IsTrue(!string.IsNullOrEmpty(folder.ParentId), "The ParentId element should not be null.");

            // If the schema validation result is true and ParentId is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6926,
                @"[In ParentId(GetHierarchy)] The ParentId element is a required child element of the Folder element in GetHierarchy command responses that specifies the server ID of the folder's parent folder. ");

            // If the schema validation result is true and ParentId is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6929,
                @"[In ParentId(GetHierarchy)] Element ParentId in GetHierarchy command response (section 2.2.2.7), the parent element is Folder (section 2.2.3.66.1).");

            // If the schema validation result is true and ParentId is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6930,
                @"[In ParentId(GetHierarchy)] None [Element ParentId in GetHierarchy command response (section 2.2.2.7) has no child element .]");

            // If the schema validation result is true and ParentId is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6931,
                @"[In ParentId(GetHierarchy)] Element ParentId in FolderUpdate command response (section 2.2.2.7), the data type is string ([MS-ASDTYPE] section 2.7).");

            // If the schema validation result is true and ParentId is not null, this requirement can be verified.
            this.Site.CaptureRequirement(
                6932,
                @"[In ParentId(GetHierarchy)] Element ParentId in GetHierarchy command response (section 2.2.2.7), the number allowed is 1…1 (required).");

            // If the schema validation result is true, this requirement can be verified.
            this.Site.CaptureRequirement(
                7384,
                @"[In Type(GetHierarchy)] Element Type in GetHierarchy command response (section 2.2.2.7), the parent element is Folder (section 2.2.3.66.1).");

            // If the schema validation result is true, this requirement can be verified.
            this.Site.CaptureRequirement(
                7385,
                @"[In Type(GetHierarchy)] None [Element Type in GetHierarchy command response (section 2.2.2.7) has no child element.]");

            // If the schema validation result is true, this requirement can be verified.
            this.Site.CaptureRequirement(
                7386,
                @"[In Type(GetHierarchy)] Element Type in GetHierarchy command response (section 2.2.2.7), the data type is unsignedByte ([MS-ASDTYPE] section 2.8).");

            // If the schema validation result is true, this requirement can be verified.
            this.Site.CaptureRequirement(
                7387,
                @"[In Type(GetHierarchy)] Element Type in GetHierarchy command response (section 2.2.2.7), the number allowed is 1…1 (required).");

            bool isVerifiedR7388 = folder.Type == 1 || folder.Type == 2 || folder.Type == 3 || folder.Type == 4 || folder.Type == 5 || folder.Type == 6;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR7388,
                7388,
                @"[In Type(GetHierarchy)] The following table lists the valid values [1-6] for this element. ");
        }
        #endregion

        #region Capture DTD [MS-ASTYPE]

        #region Capture container data type

        /// <summary>
        /// This method is used to verify the container data type related requirements.
        /// </summary>
        private void VerifyContainerDataType()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        #endregion

        #region Capture dateTime data type

        /// <summary>
        /// This method is used to verify the dateTime related requirements.
        /// </summary>
        private void VerifyDateTimeStructure()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

            // If the schema validation is successful, then MS-ASDTYPE_R12 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R12
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12,
                @"[In dateTime Data Type] It [dateTime]is declared as an element whose type attribute is set to ""dateTime"".");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R20");

            // ActiveSyncClient encoded dateTime data as inline strings, so if response is successfully returned this requirement can be verified.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R20
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                20,
                @"[In dateTime Data Type] Elements with a dateTime data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

            // If the schema validation is successful, then MS-ASDTYPE_R15 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R15
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
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

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R18");

            // If the schema validation is successful, then MS-ASDTYPE_R18 can be captured.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R18
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                18,
                @"[In dateTime Data Type] Note: Dates and times in calendar items (as specified in [MS-ASCAL]) MUST NOT include punctuation separators.");
        }

        #endregion

        #region Capture integer data type

        /// <summary>
        /// This method is used to verify the integer data type related requirements.
        /// </summary>
        private void VerifyIntegerDataType()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R86");

            // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                86,
                @"[In integer Data Type] It [an integer] is an XML schema primitive data type, as specified in [XMLSCHEMA2/2] section 3.3.13.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R87");

            // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R87
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                87,
                @"[In integer Data Type] Elements with an integer data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");
        }

        #endregion

        #region Capture string data type

        /// <summary>
        /// This method is used to verify the string data type related requirements.
        /// </summary>
        private void VerifyStringDataType()
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                90,
                @"[In string Data Type] An element of this [string] type is declared as an element with a type attribute of ""string"".");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R91");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R91
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                91,
                @"[In string Data Type] Elements with a string data type MUST be encoded and transmitted as [WBXML1.2] inline strings.");

            // ActiveSyncClient encoded string data as inline strings, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");
        }
        #endregion

        #region Capture unsignedbyte data type

        /// <summary>
        /// This method is used to verify the unsignedByte data type related requirements.
        /// </summary>
        /// <param name="byteValue">A byte value.</param>
        private void VerifyUnsignedByteDataType(byte byteValue)
        {
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R123");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R123
            Site.CaptureRequirementIfIsTrue(
                byteValue <= 255,
                "MS-ASDTYPE",
                123,
                @"[In unsignedByte Data Type] The unsignedByte data type is an integer value between 0 and 255, inclusive.");

            // Add the debug information.
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

        #endregion

        #region Capture ATD [MS-ASWBXML]

        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        /// <param name="cmdName">Current MS-ASCMD command name.</param>
        /// <param name="response">MS-ASCMD response.</param>
        private void VerifyWBXMLCapture(CommandName cmdName, object response)
        {
            #region Get the WbxmlTracer Instance for WBXML data.
            this.msaswbxmlImplementation = this.activeSyncClient.GetMSASWBXMLImplementationInstance();

            if (cmdName == CommandName.GetAttachment)
            {
                // Ignore the GetAttachment command as it does not have any WBXML process.
                return;
            }
            #endregion

            #region Capture global WBXML tokens related requirements
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R801");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                801,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 0 [for Token name] SWITCH_PAGE, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R802");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                802,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 1 [for Token name] END, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R803");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                803,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 2 [for Token name] ENTITY, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R804");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                804,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 3 [for Token name] STR_I, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R805");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                805,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 4 [for Token name] LITERAL, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R806");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                806,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 40 [for Token name] EXT_I_0, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R807");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                807,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 41 [for Token name] EXT_I_1, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R808");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                808,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 42 [for Token name] EXT_I_2, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R809");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                809,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 43 [for Token name] PI, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R810");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                810,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 44 [for Token name] LITERAL_C, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R811");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                811,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 80 [for Token name] EXT_T_0, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R812");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                812,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 81 [for Token name] EXT_T_1, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R813");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                813,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 82 [for Token name] EXT_T_2, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R814");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                814,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 83 [for Token name] STR_T, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R815");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                815,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] 84 [for Token name] LITERAL_A, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R816");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                816,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] C0 [for Token name] EXT_0, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R817");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                817,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] C1 [for Token name] EXT_1, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R818");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                818,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] C2 [for Token name] EXT_2, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R819");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                819,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] C3 [for Token name] OPAQUE, Reference [WBXML1.2].");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R820");

            // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
            Site.CaptureRequirement(
                "MS-ASWBXML",
                820,
                @"[In Standards Assignments] [This algorithm uses the global WBXML token] C4 [for Token name] LITERAL_AC, Reference [WBXML1.2].");
            #endregion

            #region Capture Code Pages related requirements
            AdapterHelper adapterHelper = new AdapterHelper();
            byte statusOfResponses = adapterHelper.GetStatusFromResponses(response);

            // Status 102 means the WBXML decode/encode error on Server
            if (statusOfResponses != 102)
            {
                // Get decode data and capture requirement for decode processing
                Dictionary<string, int> decodeData = this.msaswbxmlImplementation.DecodeDataCollection;

                foreach (KeyValuePair<string, int> decodeDataItem in decodeData)
                {
                    byte token;
                    string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                    string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                    int currentCodepage = decodeDataItem.Value;

                    bool isValidCodePage = currentCodepage >= 0 && currentCodepage <= 25;
                    Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24,actual value is :{0}", currentCodepage);

                    // begin to capture requirement
                    this.CaptureCodePageRequirement(currentCodepage, codePageName, tagName, token);

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R1");

                    // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
                    Site.CaptureRequirement(
                        "MS-ASWBXML",
                        1,
                        @"[In ActiveSync WBXML Algorithm Details] ActiveSync messages are transported as HTTP POST messages, as specified in [MS-ASHTTP], where the body of the message contains WBXML formatted data.");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R6");

                    // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
                    Site.CaptureRequirement(
                        "MS-ASWBXML",
                        6,
                        @"[In Initialization] The XML tags in both request and response messages are encoded by using WBXML tokenization, as specified in [WBXML1.2].");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R8");

                    // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
                    Site.CaptureRequirement(
                        "MS-ASWBXML",
                        8,
                        @"[In Initialization] WBXML parsers MUST use the WBXML code pages specified in the following sections [2.1.2.1   Code Pages].");

                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R655");

                    // ActiveSyncClient will encode and decode the response by using WBXML, so if response is successfully returned this requirement can be covered.
                    Site.CaptureRequirement(
                        "MS-ASWBXML",
                        655,
                        @"[In Processing Rules] This algorithm uses the following features that are specified in [WBXML1.2]: 
WBXML tokens to encode XML tags
WBXML code pages to support multiple XML namespaces
Inline strings
Opaque data");
                }
            }

            string protocolVersion = Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site);

            if (string.Compare(protocolVersion, "14.0") == 0 || string.Compare(protocolVersion, "14.1") == 0 || string.Compare(protocolVersion, "16.0") == 0)
            {
                if (this.isClassTagInPage0Exist)
                {
                    // Add the debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R825");

                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R825
                    Site.CaptureRequirementIfIsTrue(
                        this.isClassTagInPage6Exist == false,
                        "MS-ASWBXML",
                        825,
                        @"[In Code Page 6: GetItemEstimate] Note 1: The Class tag in WBXML code page 0 (AirSync) is used instead of the Class tag in WBXML code page 6 with protocol versions 14.0, 14.1, and 16.0.");
                }
            }
            #endregion
        }

        /// <summary>
        /// Verify the requirements for code pages.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="codePageName">Code page name.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureCodePageRequirement(int codePageNumber, string codePageName, string tagName, byte token)
        {
            // capture the tag and token mapping
            switch (codePageNumber)
            {
                case 0:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R10");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R10
                        Site.CaptureRequirementIfAreEqual<string>(
                            "airsync",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            10,
                            @"[In Code Pages][This algorithm supports] [Code page] 0[that indicates] [XML namespace] AirSync.");

                        this.CaptureRequirementsRelateToCodePage0(codePageNumber, tagName, token);
                        break;
                    }

                case 1:
                    {
                        this.CaptureRequirementsRelateToCodePage1(codePageNumber, tagName, token);
                        break;
                    }

                case 2:
                    {
                        this.CaptureRequirementsRelateToCodePage2(codePageNumber, tagName, token);
                        break;
                    }

                case 3:
                    {
                        break;
                    }

                case 4:
                    {
                        this.CaptureRequirementsRelateToCodePage4(codePageNumber, tagName, token);
                        break;
                    }

                case 5:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R15");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R15
                        Site.CaptureRequirementIfAreEqual<string>(
                            "move",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            15,
                            @"[In Code Pages][This algorithm supports] [Code page] 5 [that indicates][XML namespace] Move");

                        this.CaptureRequirementsRelateToCodePage5(codePageNumber, tagName, token);
                        break;
                    }

                case 6:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R16");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R16
                        Site.CaptureRequirementIfAreEqual<string>(
                            "getitemestimate",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            16,
                            @"[In Code Pages] [This algorithm supports] [Code page] 6 [that indicates] [XML namespace] GetItemEstimate");
                        this.CaptureRequirementsRelateToCodePage6(codePageNumber, tagName, token);
                        break;
                    }

                case 7:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R17");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R17
                        Site.CaptureRequirementIfAreEqual<string>(
                            "folderhierarchy",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            17,
                            @"[In Code Pages] [This algorithm supports][Code page] 7[that indicates] [XML namespace] FolderHierarchy");

                        this.CaptureRequirementsRelateToCodePage7(codePageNumber, tagName, token);
                        break;
                    }

                case 8:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R18");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R18
                        Site.CaptureRequirementIfAreEqual<string>(
                            "meetingresponse",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            18,
                            @"[In Code Pages] [This algorithm supports][Code page] 8[that indicates] [XML namespace] MeetingResponse");

                        this.CaptureRequirementsRelateToCodePage8(codePageNumber, tagName, token);
                        break;
                    }

                case 9:
                    {
                        this.CaptureRequirementsRelateToCodePage9(codePageNumber, tagName, token);
                        break;
                    }

                case 10:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R20");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R20
                        Site.CaptureRequirementIfAreEqual<string>(
                            "resolverecipients",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            20,
                            @"[In Code Pages][This algorithm supports] [Code page] 10[that indicates] [XML namespace] ResolveRecipients");

                        this.CaptureRequirementsRelateToCodePage10(codePageNumber, tagName, token);
                        break;
                    }

                case 11:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R21");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R21
                        Site.CaptureRequirementIfAreEqual<string>(
                            "validatecert",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            21,
                            @"[In Code Pages] [This algorithm supports] [Code page] 11 [that indicates] [XML namespace] ValidateCert");

                        this.CaptureRequirementsRelateToCodePage11(codePageNumber, tagName, token);
                        break;
                    }

                case 12:
                    {
                        this.CaptureRequirementsRelateToCodePage12(codePageNumber, tagName, token);
                        break;
                    }

                case 13:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R23");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R23
                        Site.CaptureRequirementIfAreEqual<string>(
                            "ping",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            23,
                            @"[In Code Pages][This algorithm supports] [Code page] 13 [that indicates][XML namespace] Ping");

                        this.CaptureRequirementsRelateToCodePage13(codePageNumber, tagName, token);
                        break;
                    }

                case 14:
                    {
                        this.CaptureRequirementsRelateToCodePage14(codePageNumber, tagName, token);
                        break;
                    }

                case 15:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R25");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R25
                        Site.CaptureRequirementIfAreEqual<string>(
                            "search",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            25,
                            @"[In Code Pages][This algorithm supports] [Code page] 15 [that indicates][XML namespace] Search");

                        this.CaptureRequirementsRelateToCodePage15(codePageNumber, tagName, token);
                        break;
                    }

                case 16:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R26");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R26
                        Site.CaptureRequirementIfAreEqual<string>(
                            "gal",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            26,
                            @"[In Code Pages] [This algorithm supports] [Code page] 16 [that indicates] [XML namespace] Gal");
                        this.CaptureRequirementsRelateToCodePage16(codePageNumber, tagName, token);
                        break;
                    }

                case 17:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R27");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R27 
                        Site.CaptureRequirementIfAreEqual<string>(
                            "airsyncbase",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            27,
                            @"[In Code Pages] [This algorithm supports][Code page] 17 [that indicates][XML namespace] AirSyncBase");

                        this.CaptureRequirementsRelateToCodePage17(codePageNumber, tagName, token);
                        break;
                    }

                case 18:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R28");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R28
                        Site.CaptureRequirementIfAreEqual<string>(
                            "settings",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            28,
                            @"[In Code Pages][This algorithm supports] [Code page] 18[that indicates] [XML namespace] Settings");

                        this.CaptureRequirementsRelateToCodePage18(codePageNumber, tagName, token);
                        break;
                    }

                case 19:
                    {
                        this.CaptureRequirementsRelateToCodePage19(codePageNumber, tagName, token);
                        break;
                    }

                case 20:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R30");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R30
                        Site.CaptureRequirementIfAreEqual<string>(
                            "itemoperations",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            30,
                            @"[In Code Pages] [This algorithm supports][Code page] 20 [that indicates][XML namespace] ItemOperations");

                        this.CaptureRequirementsRelateToCodePage20(codePageNumber, tagName, token);
                        break;
                    }

                case 21:
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R31");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R31
                        Site.CaptureRequirementIfAreEqual<string>(
                            "composemail",
                            codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            31,
                            @"[In Code Pages] [This algorithm supports][Code page] 21[that indicates] [XML namespace] ComposeMail");

                        this.CaptureRequirementsRelateToCodePage21(codePageNumber, tagName, token);
                        break;
                    }

                case 22:
                    {
                        this.CaptureRequirementsRelateToCodePage22(codePageNumber, tagName, token);
                        break;
                    }

                case 23:
                    {
                        this.CaptureRequirementsRelateToCodePage23(codePageNumber, tagName, token);
                        break;
                    }

                case 24:
                    {
                        this.CaptureRequirementsRelateToCodePage24(codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        #region tag and token mapping captures.

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 0.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage0(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Sync":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R36");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R36
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            36,
                            @"[In Code Page 0: AirSync] [Tag name] Sync [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Responses":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R37");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R37
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            37,
                            @"[In Code Page 0: AirSync] [Tag name] Responses [Token] 0x06 [supports protocol versions] All");

                        break;
                    }

                case "Add":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R38");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R38
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            38,
                            @"[In Code Page 0: AirSync] [Tag name] Add [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "Change":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R39");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R39
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            39,
                            @"[In Code Page 0: AirSync] [Tag name] Change [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Delete":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R40");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R40
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            40,
                            @"[In Code Page 0: AirSync] [Tag name] Delete [Token] 0x09 [supports protocol versions] All");

                        break;
                    }

                case "Fetch":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R41");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R41
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            41,
                            @"[In Code Page 0: AirSync] [Tag name] Fetch [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "SyncKey":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R42");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R42
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            42,
                            @"[In Code Page 0: AirSync] [Tag name] SyncKey [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "ClientId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R43");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R43
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            43,
                            @"[In Code Page 0: AirSync] [Tag name] ClientId [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "ServerId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R44");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R44
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            44,
                            @"[In Code Page 0: AirSync] [Tag name] ServerId [Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R45");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R45
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            45,
                            @"[In Code Page 0: AirSync] [Tag name] Status [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                case "Collection":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R46");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R46
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            46,
                            @"[In Code Page 0: AirSync] [Tag name] Collection [Token] 0x0F [supports protocol versions] All");

                        break;
                    }

                case "Class":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R47");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R47
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            47,
                            @"[In Code Page 0: AirSync] [Tag name] Class [Token] 0x10 [supports protocol versions] All");

                        this.isClassTagInPage0Exist = true;

                        break;
                    }

                case "CollectionId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R48");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R48
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            48,
                            @"[In Code Page 0: AirSync] [Tag name] CollectionId [Token] 0x12 [supports protocol versions] All");

                        break;
                    }

                case "GetChanges":
                    {
                        break;
                    }

                case "MoreAvailable":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R50");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R50
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            50,
                            @"[In Code Page 0: AirSync] [Tag name] MoreAvailable [Token] 0x14 [supports protocol versions] All");

                        break;
                    }

                case "WindowSize":
                    {
                        break;
                    }

                case "Commands":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R52");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R52
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            52,
                            @"[In Code Page 0: AirSync] [Tag name] Commands [Token] 0x16 [supports protocol versions] All");

                        break;
                    }

                case "Options":
                    {
                        break;
                    }

                case "FilterType":
                    {
                        break;
                    }

                case "Conflict":
                    {
                        break;
                    }

                case "Collections":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R56");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R56
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1C,
                            token,
                            "MS-ASWBXML",
                            56,
                            @"[In Code Page 0: AirSync] [Tag name] Collections [Token] 0x1C [supports protocol versions] All");

                        break;
                    }

                case "ApplicationData":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R57");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R57
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1D,
                            token,
                            "MS-ASWBXML",
                            57,
                            @"[In Code Page 0: AirSync] [Tag name] ApplicationData [Token] 0x1D [supports protocol versions] All");

                        break;
                    }

                case "DeletesAsMoves":
                    {
                        break;
                    }

                case "Supported":
                    {
                        break;
                    }

                case "SoftDelete":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R60");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R60
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x21,
                            token,
                            "MS-ASWBXML",
                            60,
                            @"[In Code Page 0: AirSync] [Tag name] SoftDelete [Token] 0x21 [supports protocol versions] All");
                        break;
                    }

                case "MIMESupport":
                    {
                        break;
                    }

                case "MIMETruncation":
                    {
                        break;
                    }

                case "Wait":
                    {
                        break;
                    }

                case "Limit":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R64");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R64
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x25,
                            token,
                            "MS-ASWBXML",
                            64,
                            @"[In Code Page 0: AirSync] [Tag name] Limit [Token] 0x25 [supports protocol versions] 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Partial":
                    {
                        break;
                    }

                case "ConversationMode":
                    {
                        break;
                    }

                case "MaxItems":
                    {
                        break;
                    }

                case "HeartbeatInterval":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 1.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage1(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Anniversary":
                    {
                        break;
                    }

                case "AssistantName":
                    {
                        break;
                    }

                case "AssistantPhoneNumber":
                    {
                        break;
                    }

                case "Birthday":
                    {
                        break;
                    }

                case "Business2PhoneNumber":
                    {
                        break;
                    }

                case "BusinessAddressCity":
                    {
                        break;
                    }

                case "BusinessAddressCountry":
                    {
                        break;
                    }

                case "BusinessAddressPostalCode":
                    {
                        break;
                    }

                case "BusinessAddressState":
                    {
                        break;
                    }

                case "BusinessAddressStreet":
                    {
                        break;
                    }

                case "BusinessFaxNumber":
                    {
                        break;
                    }

                case "BusinessPhoneNumber":
                    {
                        break;
                    }

                case "CarPhoneNumber":
                    {
                        break;
                    }

                case "Categories":
                    {
                        break;
                    }

                case "Category":
                    {
                        break;
                    }

                case "Children":
                    {
                        break;
                    }

                case "Child":
                    {
                        break;
                    }

                case "CompanyName":
                    {
                        break;
                    }

                case "Department":
                    {
                        break;
                    }

                case "Email1Address":
                    {
                        break;
                    }

                case "Email2Address":
                    {
                        break;
                    }

                case "Email3Address":
                    {
                        break;
                    }

                case "FileAs":
                    {
                        break;
                    }

                case "FirstName":
                    {
                        break;
                    }

                case "Home2PhoneNumber":
                    {
                        break;
                    }

                case "HomeAddressCity":
                    {
                        break;
                    }

                case "HomeAddressCountry":
                    {
                        break;
                    }

                case "HomeAddressPostalCode":
                    {
                        break;
                    }

                case "HomeAddressState":
                    {
                        break;
                    }

                case "HomeAddressStreet":
                    {
                        break;
                    }

                case "HomeFaxNumber":
                    {
                        break;
                    }

                case "HomePhoneNumber":
                    {
                        break;
                    }

                case "JobTitle":
                    {
                        break;
                    }

                case "LastName":
                    {
                        break;
                    }

                case "MiddleName":
                    {
                        break;
                    }

                case "MobilePhoneNumber":
                    {
                        break;
                    }

                case "OfficeLocation":
                    {
                        break;
                    }

                case "OtherAddressCity":
                    {
                        break;
                    }

                case "OtherAddressCountry":
                    {
                        break;
                    }

                case "OtherAddressPostalCode":
                    {
                        break;
                    }

                case "OtherAddressState":
                    {
                        break;
                    }

                case "OtherAddressStreet":
                    {
                        break;
                    }

                case "PagerNumber":
                    {
                        break;
                    }

                case "RadioPhoneNumber":
                    {
                        break;
                    }

                case "Spouse":
                    {
                        break;
                    }

                case "Suffix":
                    {
                        break;
                    }

                case "Title":
                    {
                        break;
                    }

                case "WebPage":
                    {
                        break;
                    }

                case "YomiCompanyName":
                    {
                        break;
                    }

                case "YomiFirstName":
                    {
                        break;
                    }

                case "YomiLastName":
                    {
                        break;
                    }

                case "Picture":
                    {
                        break;
                    }

                case "Alias":
                    {
                        break;
                    }

                case "WeightedRank":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 2.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage2(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "DateReceived":
                    {
                        break;
                    }

                case "DisplayTo":
                    {
                        break;
                    }

                case "Importance":
                    {
                        break;
                    }

                case "MessageClass":
                    {
                        break;
                    }

                case "Subject":
                    {
                        break;
                    }

                case "Read":
                    {
                        break;
                    }

                case "To":
                    {
                        break;
                    }

                case "Cc":
                    {
                        break;
                    }

                case "From":
                    {
                        break;
                    }

                case "ReplyTo":
                    {
                        break;
                    }

                case "AllDayEvent":
                    {
                        break;
                    }

                case "Categories":
                    {
                        break;
                    }

                case "Category":
                    {
                        break;
                    }

                case "DtStamp":
                    {
                        break;
                    }

                case "EndTime":
                    {
                        break;
                    }

                case "InstanceType":
                    {
                        break;
                    }

                case "BusyStatus":
                    {
                        break;
                    }

                case "Location":
                    {
                        break;
                    }

                case "MeetingRequest":
                    {
                        break;
                    }

                case "Organizer":
                    {
                        break;
                    }

                case "RecurrenceId":
                    {
                        break;
                    }

                case "Reminder":
                    {
                        break;
                    }

                case "ResponseRequested":
                    {
                        break;
                    }

                case "Recurrences":
                    {
                        break;
                    }

                case "Recurrence":
                    {
                        break;
                    }

                case "Type":
                    {
                        break;
                    }

                case "Until":
                    {
                        break;
                    }

                case "Occurrences":
                    {
                        break;
                    }

                case "Interval":
                    {
                        break;
                    }

                case "DayOfWeek":
                    {
                        break;
                    }

                case "DayOfMonth":
                    {
                        break;
                    }

                case "WeekOfMonth":
                    {
                        break;
                    }

                case "MonthOfYear":
                    {
                        break;
                    }

                case "StartTime":
                    {
                        break;
                    }

                case "Sensitivity":
                    {
                        break;
                    }

                case "TimeZone":
                    {
                        break;
                    }

                case "GlobalObjId":
                    {
                        break;
                    }

                case "ThreadTopic":
                    {
                        break;
                    }

                case "InternetCPID":
                    {
                        break;
                    }

                case "Flag":
                    {
                        break;
                    }

                case "Status":
                    {
                        break;
                    }

                case "ContentClass":
                    {
                        break;
                    }

                case "FlagType":
                    {
                        break;
                    }

                case "CompleteTime":
                    {
                        break;
                    }

                case "DisallowNewTimeProposal":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 4.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage4(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Timezone":
                    {
                        break;
                    }

                case "AllDayEvent":
                    {
                        break;
                    }

                case "Attendees":
                    {
                        break;
                    }

                case "Attendee":
                    {
                        break;
                    }

                case "Email":
                    {
                        break;
                    }

                case "Name":
                    {
                        break;
                    }

                case "BusyStatus":
                    {
                        break;
                    }

                case "Categories":
                    {
                        break;
                    }

                case "Category":
                    {
                        break;
                    }

                case "DtStamp":
                    {
                        break;
                    }

                case "EndTime":
                    {
                        break;
                    }

                case "Exception":
                    {
                        break;
                    }

                case "Exceptions":
                    {
                        break;
                    }

                case "Deleted":
                    {
                        break;
                    }

                case "ExceptionStartTime":
                    {
                        break;
                    }

                case "Location":
                    {
                        break;
                    }

                case "MeetingStatus":
                    {
                        break;
                    }

                case "OrganizerEmail":
                    {
                        break;
                    }

                case "OrganizerName":
                    {
                        break;
                    }

                case "Recurrence":
                    {
                        break;
                    }

                case "Type":
                    {
                        break;
                    }

                case "Until":
                    {
                        break;
                    }

                case "Occurrences":
                    {
                        break;
                    }

                case "Interval":
                    {
                        break;
                    }

                case "DayOfWeek":
                    {
                        break;
                    }

                case "DayOfMonth":
                    {
                        break;
                    }

                case "WeekOfMonth":
                    {
                        break;
                    }

                case "MonthOfYear":
                    {
                        break;
                    }

                case "Reminder":
                    {
                        break;
                    }

                case "Sensitivity":
                    {
                        break;
                    }

                case "Subject":
                    {
                        break;
                    }

                case "StartTime":
                    {
                        break;
                    }

                case "UID":
                    {
                        break;
                    }

                case "AttendeeStatus":
                    {
                        break;
                    }

                case "AttendeeType":
                    {
                        break;
                    }

                case "DisallowNewTimeProposal":
                    {
                        break;
                    }

                case "ResponseRequested":
                    {
                        break;
                    }

                case "AppointmentReplyTime":
                    {
                        break;
                    }

                case "ResponseType":
                    {
                        break;
                    }

                case "CalendarType":
                    {
                        break;
                    }

                case "IsLeapMonth":
                    {
                        break;
                    }

                case "FirstDayOfWeek":
                    {
                        break;
                    }

                case "OnlineMeetingConfLink":
                    {
                        break;
                    }

                case "OnlineMeetingExternalLink":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 5.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage5(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "MoveItems":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R234");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R234
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            234,
                            @"[In Code Page 5: Move] [Tag name] MoveItems [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Move":
                    {
                        break;
                    }

                case "SrcMsgId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R236");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R236
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            236,
                            @"[In Code Page 5: Move] [Tag name] SrcMsgId [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "SrcFldId":
                    {
                        break;
                    }

                case "DstFldId":
                    {
                        break;
                    }

                case "Response":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R239");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R239
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            239,
                            @"[In Code Page 5: Move] [Tag name] Response [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R240");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R240
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            240,
                            @"[In Code Page 5: Move] [Tag name] Status [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "DstMsgId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R241");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R241
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            241,
                            @"[In Code Page 5: Move] [Tag name] DstMsgId [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 6.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage6(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "GetItemEstimate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R243");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R243
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            243,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] GetItemEstimate [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Version":
                    {
                        break;
                    }

                case "Collections":
                    {
                        break;
                    }

                case "Collection":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R246");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R246
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            246,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] Collection [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Class":
                    {
                        this.isClassTagInPage6Exist = true;
                        break;
                    }

                case "CollectionId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R248");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R248
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            248,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] CollectionId [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "DateTime":
                    {
                        break;
                    }

                case "Estimate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R250");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R250
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            250,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] Estimate [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "Response":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R251");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R251
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            251,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] Response [Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R252");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R252
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            252,
                            @"[In Code Page 6: GetItemEstimate] [Tag name] Status [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 7.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage7(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "DisplayName":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R254");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R254
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            254,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] DisplayName [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "ServerId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R255");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R255
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            255,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] ServerId [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "ParentId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R260");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R260
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            260,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] ParentId [Token] 0x09 [supports protocol versions] All");

                        break;
                    }

                case "Type":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R261");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R261
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0A,
                            token,
                            "MS-ASWBXML",
                            261,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Type [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R262");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R262
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            262,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Status  [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "Changes":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R263");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R263
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            263,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Changes [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                case "Add":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R264");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R264
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            264,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Add [Token] 0x0F [supports protocol versions] All");

                        break;
                    }

                case "Delete":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R265");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R265
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            265,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Delete [Token] 0x10 [supports protocol versions] All");

                        break;
                    }

                case "Update":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R266");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R266
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x11,
                            token,
                             "MS-ASWBXML",
                            266,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Update [Token] 0x11 [supports protocol versions] All");

                        break;
                    }

                case "SyncKey":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R267");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R267
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            267,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] SyncKey [Token] 0x12 [supports protocol versions] All");

                        break;
                    }

                case "FolderCreate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R268");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R268
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            268,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] FolderCreate [Token] 0x13 [supports protocol versions] All");

                        break;
                    }

                case "FolderDelete":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R269");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R269
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            269,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] FolderDelete [Token] 0x14 [supports protocol versions] All");

                        break;
                    }

                case "FolderUpdate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R270");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R270
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x15,
                            token,
                            "MS-ASWBXML",
                            270,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] FolderUpdate [Token] 0x15 [supports protocol versions] All");

                        break;
                    }

                case "FolderSync":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R271");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R271
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            271,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] FolderSync [Token] 0x16 [supports protocol versions] All");

                        break;
                    }

                case "Count":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R272");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R272
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x17,
                            token,
                            "MS-ASWBXML",
                            272,
                            @"[In Code Page 7: FolderHierarchy] [Tag name] Count [Token] 0x17 [supports protocol versions] All");

                        break;
                    }

                case "Folders":
                    {
                         break;
                    }

                case "Folder":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 8.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage8(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "CalendarId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R273");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R273
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            273,
                            @"[In Code Page 8: MeetingResponse] [Tag name] CalendarId [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "CollectionId":
                    {
                        break;
                    }

                case "MeetingResponse":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R275");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R275
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x07,
                            token,
                            "MS-ASWBXML",
                            275,
                            @"[In Code Page 8: MeetingResponse] [Tag name] MeetingResponse [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "RequestId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R276");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R276
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            276,
                            @"[In Code Page 8: MeetingResponse] [Tag name] RequestId [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Request":
                    {
                        break;
                    }

                case "Result":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R278");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R278
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            278,
                            @"[In Code Page 8: MeetingResponse] [Tag name] Result [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R279");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R279
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            279,
                            @"[In Code Page 8: MeetingResponse] [Tag name] Status [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "UserResponse":
                    {
                        break;
                    }

                case "InstanceId":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 9.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage9(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Categories":
                    {
                        break;
                    }

                case "Category":
                    {
                        break;
                    }

                case "Complete":
                    {
                        break;
                    }

                case "DateCompleted":
                    {
                        break;
                    }

                case "DueDate":
                    {
                        break;
                    }

                case "UtcDueDate":
                    {
                        break;
                    }

                case "Importance":
                    {
                        break;
                    }

                case "Recurrence":
                    {
                        break;
                    }

                case "ReminderSet":
                    {
                        break;
                    }

                case "ReminderTime":
                    {
                        break;
                    }

                case "Sensitivity":
                    {
                        break;
                    }

                case "StartDate":
                    {
                        break;
                    }

                case "UtcStartDate":
                    {
                        break;
                    }

                case "Subject":
                    {
                        break;
                    }

                case "OrdinalDate":
                    {
                        break;
                    }

                case "SubOrdinalDate":
                    {
                        break;
                    }

                case "CalendarType":
                    {
                        break;
                    }

                case "IsLeapMonth":
                    {
                        break;
                    }

                case "FirstDayOfWeek":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 10.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage10(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "ResolveRecipients":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R312");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R312
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            312,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] ResolveRecipients [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Response":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R313");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R313
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            313,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Response [Token] 0x06 [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R314");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R314
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            314,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Status [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "Type":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R315");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R315
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            315,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Type [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Recipient":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R316");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R316
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            316,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Recipient  [Token] 0x09 [supports protocol versions] All");

                        break;
                    }

                case "DisplayName":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R317");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R317
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0A,
                            token,
                            "MS-ASWBXML",
                            317,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] DisplayName [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "EmailAddress":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R318");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R318
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            318,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] EmailAddress [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "Certificates":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R319");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R319
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0C,
                            token,
                            "MS-ASWBXML",
                            319,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Certificates [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "Certificate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R320");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R320
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            320,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] Certificate [Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "MiniCertificate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R321");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R321
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            321,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] MiniCertificate [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                case "Options":
                    {
                        break;
                    }

                case "To":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R323");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R323
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            323,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] To [Token] 0x10 [supports protocol versions] All");

                        break;
                    }

                case "CertificateRetrieval":
                    {
                        break;
                    }

                case "RecipientCount":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R325");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R325
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            325,
                            @"[In Code Page 10: ResolveRecipients] [Tag name] RecipientCount [Token] 0x12 [supports protocol versions] All");

                        break;
                    }

                case "MaxCertificates":
                    {
                        break;
                    }

                case "MaxAmbiguousRecipients":
                    {
                        break;
                    }

                case "CertificateCount":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R328");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R328
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x15,
                                token,
                                "MS-ASWBXML",
                                328,
                                @"[In Code Page 10: ResolveRecipients] [Tag name] CertificateCount [Token] 0x15 [supports protocol versions] All");
                        }

                        break;
                    }

                case "Availability":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R329");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R329
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x16,
                                token,
                                "MS-ASWBXML",
                                329,
                                @"[In Code Page 10: ResolveRecipients] [Tag name] Availability [Token] 0x16 [supports protocol versions] 14.0, 14.1, 16.0");
                        }

                        break;
                    }

                case "StartTime":
                    {
                        break;
                    }

                case "EndTime":
                    {
                        break;
                    }

                case "MergedFreeBusy":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R332");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R332
                            Site.CaptureRequirementIfAreEqual<byte>(
                                 0x19,
                            token,
                                "MS-ASWBXML",
                                332,
                                @"[In Code Page 10: ResolveRecipients] [Tag name] MergedFreeBusy [Token] 0x19 [supports protocol versions] 14.0, 14.1, 16.0");
                        }

                        break;
                    }

                case "Picture":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R333");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R333
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x1A,
                                token,
                                "MS-ASWBXML",
                                333,
                                @"[In Code Page 10: ResolveRecipients] [Tag name] Picture [Token] 0x1A [supports protocol versions] 14.1, 16.0");
                        }

                        break;
                    }

                case "MaxSize":
                    {
                        break;
                    }

                case "Data":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R335");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R335
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x1C,
                                token,
                                "MS-ASWBXML",
                                335,
                                @"[In Code Page 10: ResolveRecipients] [Tag name] Data [Token] 0x1C [supports protocol versions] 14.1, 16.0");
                        }

                        break;
                    }

                case "MaxPictures":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 11.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage11(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "ValidateCert":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R524");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R524
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            524,
                            @"[In Code Page 11: ValidateCert] [Tag name] ValidateCert [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Certificates":
                    {
                        break;
                    }

                case "Certificate":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R526");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R526
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            526,
                            @"[In Code Page 11: ValidateCert] [Tag name] Certificate [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "CertificateChain":
                    {
                        break;
                    }

                case "CheckCRL":
                    {
                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R529");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R529
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            529,
                            @"[In Code Page 11: ValidateCert] [Tag name] Status [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 12.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage12(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "CustomerId":
                    {
                        break;
                    }

                case "GovernmentId":
                    {
                        break;
                    }

                case "IMAddress":
                    {
                        break;
                    }

                case "IMAddress2":
                    {
                        break;
                    }

                case "IMAddress3":
                    {
                        break;
                    }

                case "ManagerName":
                    {
                        break;
                    }

                case "CompanyMainPhone":
                    {
                        break;
                    }

                case "AccountName":
                    {
                        break;
                    }

                case "NickName":
                    {
                        break;
                    }

                case "MMS":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 13.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage13(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Ping":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R543");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R543
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            543,
                            @"[In Code Page 13: Ping] [Tag name] Ping [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "AutdState":
                    {
                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R545");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R545
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            545,
                            @"[In Code Page 13: Ping] [Tag name] Status [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "HeartbeatInterval":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R546");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R546
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x08,
                            token,
                            "MS-ASWBXML",
                            546,
                           @"[In Code Page 13: Ping] [Tag name] HeartbeatInterval [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Folders":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R547");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R547
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            547,
                            @"[In Code Page 13: Ping] [Tag name] Folders [Token] 0x09 [supports protocol versions] All");

                        break;
                    }

                case "Folder":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R548");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R548
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0A,
                            token,
                            "MS-ASWBXML",
                            548,
                            @"[In Code Page 13: Ping] [Tag name] Folder [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "Id":
                    {
                        break;
                    }

                case "Class":
                    {
                        break;
                    }

                case "MaxFolders":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 14.
        /// </summary>
        /// <param name="codePageNumber">code page number 14</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage14(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Provision":
                    {
                        break;
                    }

                case "Policies":
                    {
                        break;
                    }

                case "Policy":
                    {
                        break;
                    }

                case "PolicyType":
                    {
                        break;
                    }

                case "PolicyKey":
                    {
                        break;
                    }

                case "Data":
                    {
                        break;
                    }

                case "Status":
                    {
                        break;
                    }

                case "RemoteWipe":
                    {
                        break;
                    }

                case "EASProvisionDoc":
                    {
                        break;
                    }

                case "DevicePasswordEnabled":
                    {
                        break;
                    }

                case "AlphanumericDevicePasswordRequired":
                    {
                        break;
                    }

                case "DeviceEncryptionEnabled":
                    {
                        break;
                    }

                case "RequireStorageCardEncryption":
                    {
                        break;
                    }

                case "PasswordRecoveryEnabled":
                    {
                        break;
                    }

                case "AttachmentsEnabled":
                    {
                        break;
                    }

                case "MinDevicePasswordLength":
                    {
                        break;
                    }

                case "MaxInactivityTimeDeviceLock":
                    {
                        break;
                    }

                case "MaxDevicePasswordFailedAttempts":
                    {
                        break;
                    }

                case "MaxAttachmentSize":
                    {
                        break;
                    }

                case "AllowSimpleDevicePassword":
                    {
                        break;
                    }

                case "DevicePasswordExpiration":
                    {
                        break;
                    }

                case "DevicePasswordHistory":
                    {
                        break;
                    }

                case "AllowStorageCard":
                    {
                        break;
                    }

                case "AllowCamera":
                    {
                        break;
                    }

                case "RequireDeviceEncryption":
                    {
                        break;
                    }

                case "AllowUnsignedApplications":
                    {
                        break;
                    }

                case "AllowUnsignedInstallationPackages":
                    {
                        break;
                    }

                case "MinDevicePasswordComplexCharacters":
                    {
                        break;
                    }

                case "AllowWiFi":
                    {
                        break;
                    }

                case "AllowTextMessaging":
                    {
                        break;
                    }

                case "AllowPOPIMAPEmail":
                    {
                        break;
                    }

                case "AllowBluetooth":
                    {
                        break;
                    }

                case "AllowIrDA":
                    {
                        break;
                    }

                case "RequireManualSyncWhenRoaming":
                    {
                        break;
                    }

                case "AllowDesktopSync":
                    {
                        break;
                    }

                case "MaxCalendarAgeFilter":
                    {
                        break;
                    }

                case "AllowHTMLEmail":
                    {
                        break;
                    }

                case "MaxEmailAgeFilter":
                    {
                        break;
                    }

                case "MaxEmailBodyTruncationSize":
                    {
                        break;
                    }

                case "MaxEmailHTMLBodyTruncationSize":
                    {
                        break;
                    }

                case "RequireSignedSMIMEMessages":
                    {
                        break;
                    }

                case "RequireEncryptedSMIMEMessages":
                    {
                        break;
                    }

                case "RequireSignedSMIMEAlgorithm":
                    {
                        break;
                    }

                case "RequireEncryptionSMIMEAlgorithm":
                    {
                        break;
                    }

                case "AllowSMIMEEncryptionAlgorithmNegotiation":
                    {
                        break;
                    }

                case "AllowSMIMESoftCerts":
                    {
                        break;
                    }

                case "AllowBrowser":
                    {
                        break;
                    }

                case "AllowConsumerEmail":
                    {
                        break;
                    }

                case "AllowRemoteDesktop":
                    {
                        break;
                    }

                case "AllowInternetSharing":
                    {
                        break;
                    }

                case "UnapprovedInROMApplicationList":
                    {
                        break;
                    }

                case "ApplicationName":
                    {
                        break;
                    }

                case "ApprovedApplicationList":
                    {
                        break;
                    }

                case "Hash":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 15.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage15(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Search":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R397");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R397
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            397,
                            @"[In Code Page 15: Search] [Tag name] Search [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Store":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R398");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R398
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            398,
                            @"[In Code Page 15: Search] [Tag name] Store [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "Name":
                    {
                        break;
                    }

                case "Query":
                    {
                        break;
                    }

                case "Options":
                    {
                        break;
                    }

                case "Range":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R402");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R402
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            402,
                            @"[In Code Page 15: Search] [Tag name] Range [Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R403");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R403
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            403,
                            @"[In Code Page 15: Search] [Tag name] Status [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "Response":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R404");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R404
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            404,
                            @"[In Code Page 15: Search] [Tag name] Response [Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "Result":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R405");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R405
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            405,
                            @"[In Code Page 15: Search] [Tag name] Result [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                case "Properties":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R406");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R406
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            406,
                            @"[In Code Page 15: Search] [Tag name] Properties [Token] 0x0F [supports protocol versions] All");

                        break;
                    }

                case "Total":
                    {
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R407");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R407
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            407,
                            @"[In Code Page 15: Search] [Tag name] Total [Token] 0x10 [supports protocol versions] All");

                        break;
                    }

                case "EqualTo":
                    {
                        break;
                    }

                case "Value":
                    {
                        break;
                    }

                case "And":
                    {
                        break;
                    }

                case "Or":
                    {
                        break;
                    }

                case "FreeText":
                    {
                        break;
                    }

                case "DeepTraversal":
                    {
                        break;
                    }

                case "LongId":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R415");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R415
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x18,
                            token,
                            "MS-ASWBXML",
                            415,
                            @"[In Code Page 15: Search] [Tag name] LongId [Token] 0x18 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "RebuildResults":
                    {
                        break;
                    }

                case "LessThan":
                    {
                        break;
                    }

                case "GreaterThan":
                    {
                        break;
                    }

                case "UserName":
                    {
                        break;
                    }

                case "Password":
                    {
                        break;
                    }

                case "ConversationId":
                    {
                        break;
                    }

                case "Picture":
                    {
                        break;
                    }

                case "MaxSize":
                    {
                        break;
                    }

                case "MaxPictures":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 16.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage16(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "DisplayName":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R430");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R430
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            430,
                            @"[In Code Page 16: GAL] [Tag name] DisplayName [Token] 0x05 [supports protocol versions] All");

                        break;
                    }

                case "Phone":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R431");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R431
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            431,
                            @"[In Code Page] [In Code Page 16: GAL] [Tag name] Phone [Token] 0x06 [supports protocol versions] All");

                        break;
                    }

                case "Office":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R432");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R432
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            432,
                            @"[In Code Page 16: GAL] [Tag name] Office [Token] 0x07 [supports protocol versions] All");

                        break;
                    }

                case "Title":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R433");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R433
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            433,
                            @"[In Code Page 16: GAL] [Tag name] Title [Token] 0x08 [supports protocol versions] All");

                        break;
                    }

                case "Company":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R434");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R434
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            434,
                            @"[In Code Page 16: GAL] [Tag name] Company[Token] 0x09 [supports protocol versions] All");

                        break;
                    }

                case "Alias":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R435");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R435
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            435,
                            @"[In Code Page 16: GAL] [Tag name] Alias [Token] 0x0A [supports protocol versions] All");

                        break;
                    }

                case "FirstName":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R436");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R436
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            436,
                            @"[In Code Page 16: GAL] [Tag name]FirstName[Token] 0x0B [supports protocol versions] All");

                        break;
                    }

                case "LastName":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R437");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R437
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            437,
                            @"[In Code Page 16: GAL] [Tag name] LastName [Token] 0x0C [supports protocol versions] All");

                        break;
                    }

                case "HomePhone":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R438");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R438
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            438,
                            @"[In Code Page 16: GAL] [Tag name] HomePhone[Token] 0x0D [supports protocol versions] All");

                        break;
                    }

                case "MobilePhone":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R439");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R439
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            439,
                            @"[In Code Page 16: GAL] [Tag name] MobilePhone [Token] 0x0E [supports protocol versions] All");

                        break;
                    }

                case "EmailAddress":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R440");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R440
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            440,
                            @"[In Code Page 16: GAL] [Tag name] EmailAddress [Token] 0x0F [supports protocol versions] All");

                        break;
                    }

                case "Picture":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R441");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R441
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            441,
                            @"[In Code Page 16: GAL] [Tag name] Picture [Token] 0x10 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R442");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R442
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            442,
                            @"[In Code Page 16: GAL] [Tag name] Status [Token] 0x11 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "Data":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R443");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R443
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            443,
                            @"[In Code Page 16: GAL] [Tag name] Data [Token] 0x12 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 17.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage17(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "BodyPreference":
                    {
                        break;
                    }

                case "Type":
                    {
                        break;
                    }

                case "TruncationSize":
                    {
                        break;
                    }

                case "AllOrNone":
                    {
                        break;
                    }

                case "Body":
                    {
                        break;
                    }

                case "Data":
                    {
                        break;
                    }

                case "EstimatedDataSize":
                    {
                        break;
                    }

                case "Truncated":
                    {
                        break;
                    }

                case "Attachments":
                    {
                        break;
                    }

                case "Attachment":
                    {
                        break;
                    }

                case "DisplayName":
                    {
                        break;
                    }

                case "FileReference":
                    {
                        break;
                    }

                case "Method":
                    {
                        break;
                    }

                case "ContentId":
                    {
                        break;
                    }

                case "ContentLocation":
                    {
                        break;
                    }

                case "IsInline":
                    {
                        break;
                    }

                case "NativeBodyType":
                    {
                        break;
                    }

                case "ContentType":
                    {
                        break;
                    }

                case "Preview":
                    {
                        break;
                    }

                case "BodyPartPreference":
                    {
                        break;
                    }

                case "BodyPart":
                    {
                        break;
                    }

                case "Status":
                    {
                        break;
                    }

                case "Location":
                    {
                        break;
                    }

                case "InstanceId":
                    {
                        break;
                    }
                case "LocationUri":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 18.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage18(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Settings":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R475");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R475
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            475,
                            @"[In Code Page 18: Settings] [Tag name] Settings [Token] 0x05 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R476");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R476
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            476,
                            @"[In Code Page 18: Settings] [Tag name] Status [Token] 0x06 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Get":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R477");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R477
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            477,
                            @"[In Code Page 18: Settings] [Tag name] Get [Token] 0x07 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Set":
                    {
                        break;
                    }

                case "Oof":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R479");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R479
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            479,
                            @"[In Code Page 18: Settings] [Tag name] Oof [Token] 0x09 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "OofState":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R480");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R480
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            480,
                            @"[In Code Page 18: Settings] [Tag name] OofState [Token] 0x0A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "StartTime":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R481");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R481
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0B,
                            token,
                            "MS-ASWBXML",
                            481,
                            @"[In Code Page 18: Settings] [Tag name] StartTime [Token] 0x0B [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "EndTime":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R482");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R482
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            482,
                            @"[In Code Page 18: Settings] [Tag name] EndTime [Token] 0x0C [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "OofMessage":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R483");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R483
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            483,
                            @"[In Code Page 18: Settings] [Tag name] OofMessage [Token] 0x0D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "AppliesToInternal":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R484");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R484
                        Site.CaptureRequirementIfAreEqual<byte>(
                             0x0E,
                            token,
                            "MS-ASWBXML",
                            484,
                            @"[In Code Page 18: Settings] [Tag name] AppliesToInternal [Token] 0x0E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "AppliesToExternalKnown":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R485");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R485
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            485,
                            @"[In Code Page 18: Settings] [Tag name] AppliesToExternalKnown [Token] 0x0F [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "AppliesToExternalUnknown":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R486");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R486
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            486,
                            @"[In Code Page 18: Settings] [Tag name] AppliesToExternalUnknown [Token] 0x10 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Enabled":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R487");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R487
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            487,
                            @"[In Code Page 18: Settings] [Tag name] Enabled [Token] 0x11 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "ReplyMessage":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R488");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R488
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            488,
                            @"[In Code Page 18: Settings] [Tag name] ReplyMessage [Token] 0x12 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "BodyType":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R489");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R489
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            489,
                            @"[In Code Page 18: Settings] [Tag name] BodyType [Token] 0x13 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "DevicePassword":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R490");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R490
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            490,
                            @"[In Code Page 18: Settings] [Tag name] DevicePassword [Token] 0x14 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Password":
                    {
                        break;
                    }

                case "Model":
                    {
                        break;
                    }

                case "IMEI":
                    {
                        break;
                    }

                case "FriendlyName":
                    {
                        break;
                    }

                case "OS":
                    {
                        break;
                    }

                case "OSLanguage":
                    {
                        break;
                    }

                case "PhoneNumber":
                    {
                        break;
                    }

                case "UserInformation":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R499");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R499
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1D,
                            token,
                            "MS-ASWBXML",
                            499,
                            @"[In Code Page 18: Settings] [Tag name] UserInformation [Token] 0x1D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "EmailAddresses":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R500");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R500
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1E,
                            token,
                            "MS-ASWBXML",
                            500,
                            @"[In Code Page 18: Settings] [Tag name] EmailAddresses [Token] 0x1E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "UserAgent":
                    {
                        break;
                    }

                case "EnableOutboundSMS":
                    {
                        break;
                    }

                case "MobileOperator":
                    {
                        break;
                    }

                case "PrimarySmtpAddress":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R505");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R505
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x23,
                            token,
                            "MS-ASWBXML",
                            505,
                            @"[In Code Page 18: Settings] [Tag name] PrimarySmtpAddress [Token]0x23 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "Accounts":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R506");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R506
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x24,
                            token,
                            "MS-ASWBXML",
                            506,
                            @"[In Code Page 18: Settings] [Tag name] Accounts [Token] 0x24 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "Account":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R507");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R507
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x25,
                            token,
                            "MS-ASWBXML",
                            507,
                            @"[In Code Page 18: Settings] [Tag name] Account [Token] 0x25 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "AccountId":
                    {
                        break;
                    }

                case "AccountName":
                    {
                        break;
                    }

                case "UserDisplayName":
                    {
                        break;
                    }

                case "SendDisabled":
                    {
                        break;
                    }

                case "DeviceInformation":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R492");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R492
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            492,
                            @"[In Code Page 18: Settings] [Tag name] DeviceInformation [Token] 0x16 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "SMTPAddress":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R501");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R501
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x1F,
                            token,
                            "MS-ASWBXML",
                            501,
                            @"[In Code Page 18: Settings] [Tag name] SMTPAddress [Token] 0x1F [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "RightsManagementInformation":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") && !Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R512");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R512
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x2B,
                                token,
                                "MS-ASWBXML",
                                512,
                                @"[In Code Page 18: Settings] [Tag name] RightsManagementInformation [Token] 0x2B [supports protocol versions] 14.1, 16.0");
                        }

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 19.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage19(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "LinkId":
                    {
                        break;
                    }

                case "DisplayName":
                    {
                        break;
                    }

                case "IsFolder":
                    {
                        break;
                    }

                case "CreationDate":
                    {
                        break;
                    }

                case "LastModifiedDate":
                    {
                        break;
                    }

                case "IsHidden":
                    {
                        break;
                    }

                case "ContentLength":
                    {
                        break;
                    }

                case "ContentType":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 20.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage20(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "ItemOperations":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R563");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R563
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            563,
                            @"[In Code Page 20: ItemOperations] [Tag name] ItemOperations [Token] 0x05 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Fetch":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R564");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R564
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            564,
                            @"[In Code Page 20: ItemOperations] [Tag name] Fetch [Token]0x06 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Store":
                    {
                        break;
                    }

                case "Options":
                    {
                        break;
                    }

                case "Range":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R567");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R567
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            567,
                            @"[In Code Page 20: ItemOperations] [Tag name] Range [Token] 0x09 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Total":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R568");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R568
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            568,
                            @"[In Code Page 20: ItemOperations] [Tag name] Total [Token] 0x0A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Properties":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R569");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R569
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            569,
                            @"[In Code Page 20: ItemOperations] [Tag name] Properties [Token] 0x0B [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Data":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R570");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R570
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            570,
                            @"[In Code Page 20: ItemOperations] [Tag name] Data [Token] 0x0C [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R571");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R571
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            571,
                            @"[In Code Page 20: ItemOperations] [Tag name] Status [Token] 0x0D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Response":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R572");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R572
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            572,
                            @"[In Code Page 20: ItemOperations] [Tag name] Response [Token] 0x0E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Version":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R573");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R573
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            573,
                            @"[In Code Page 20: ItemOperations] [Tag name] Version [Token] 0x0F [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "Schema":
                    {
                        break;
                    }

                case "Part":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R575");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R575
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            575,
                            @"[In Code Page 20: ItemOperations] [Tag name] Part [Token] 0x11 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "EmptyFolderContents":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R576");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R576
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            576,
                            @"[In Code Page 20: ItemOperations] [Tag name] EmptyFolderContents [Token] 0x12 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0");

                        break;
                    }

                case "DeleteSubFolders":
                    {
                        break;
                    }

                case "UserName":
                    {
                        break;
                    }

                case "Password":
                    {
                        break;
                    }

                case "Move":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R580");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R580
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x16,
                            token,
                                "MS-ASWBXML",
                                580,
                                @"[In Code Page 20: ItemOperations] [Tag name] Move [Token] 0x16 [supports protocol versions] 14.0, 14.1, 16.0");
                        }

                        break;
                    }

                case "DstFldId":
                    {
                        break;
                    }

                case "ConversationId":
                    {
                        if (!Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1"))
                        {
                            // Add the debug information.
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R582");

                            // Verify MS-ASWBXML requirement: MS-ASWBXML_R582
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x18,
                            token,
                                "MS-ASWBXML",
                                582,
                                @"[In Code Page 20: ItemOperations] [Tag name] ConversationId [Token] 0x18 [supports protocol versions] 14.0, 14.1, 16.0");
                        }

                        break;
                    }

                case "MoveAlways":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 21.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage21(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "SendMail":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R589");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R589
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x05,
                            token,
                            "MS-ASWBXML",
                            589,
                            @"[In Code Page 21: ComposeMail] [Tag name] SendMail [Token] 0x05 [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "SmartForward":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R590");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R590
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            590,
                            @"[In Code Page 21: ComposeMail] [Tag name] SmartForward [Token] 0x06 [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "SmartReply":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R591");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R591
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            591,
                            @"[In Code Page 21: ComposeMail] [Tag name] SmartReply [Token] 0x07 [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "SaveInSentItems":
                    {
                        break;
                    }

                case "ReplaceMime":
                    {
                        break;
                    }

                case "Source":
                    {
                        break;
                    }

                case "FolderId":
                    {
                        break;
                    }

                case "ItemId":
                    {
                        break;
                    }

                case "LongId":
                    {
                        break;
                    }

                case "InstanceId":
                    {
                        break;
                    }

                case "Mime":
                    {
                        break;
                    }

                case "ClientId":
                    {
                        break;
                    }

                case "Status":
                    {
                        // Add the debug information.
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R601");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R601
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            601,
                            @"[In Code Page 21: ComposeMail] [Tag name] Status [Token] 0x12 [supports protocol versions] 14.0, 14.1, 16.0");

                        break;
                    }

                case "AccountId":
                    {
                        break;
                    }

                case "Forwardees":
                    {
                        break;
                    }

                case "Forwardee":
                    {
                        break;
                    }

                case "ForwardeeName":
                    {
                        break;
                    }

                case "ForwardeeEmail":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 22.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage22(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "UmCallerID":
                    {
                        break;
                    }

                case "UmUserNotes":
                    {
                        break;
                    }

                case "UmAttDuration":
                    {
                        break;
                    }

                case "UmAttOrder":
                    {
                        break;
                    }

                case "ConversationId":
                    {
                        break;
                    }

                case "ConversationIndex":
                    {
                        break;
                    }

                case "LastVerbExecuted":
                    {
                        break;
                    }

                case "LastVerbExecutionTime":
                    {
                        break;
                    }

                case "ReceivedAsBcc":
                    {
                        break;
                    }

                case "Sender":
                    {
                        break;
                    }

                case "CalendarType":
                    {
                        break;
                    }

                case "IsLeapMonth":
                    {
                        break;
                    }

                case "AccountId":
                    {
                        break;
                    }

                case "FirstDayOfWeek":
                    {
                        break;
                    }

                case "MeetingMessageType":
                    {
                        break;
                    }

                case "Bcc":
                    {
                        break;
                    }

                case "IsDraft":
                    {
                        break;
                    }

                case "Send":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 23.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage23(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "Subject":
                    {
                        break;
                    }

                case "MessageClass":
                    {
                        break;
                    }

                case "LastModifiedDate":
                    {
                        break;
                    }

                case "Categories":
                    {
                        break;
                    }

                case "Category":
                    {
                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codePageNumber, tagName, token);
                        break;
                    }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 24.
        /// </summary>
        /// <param name="codePageNumber">Code page number.</param>
        /// <param name="tagName">Tag name that needs to be verified.</param>
        /// <param name="token">Token that needs to be verified.</param>
        private void CaptureRequirementsRelateToCodePage24(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "RightsManagementSupport":
                    {
                        break;
                    }

                case "RightsManagementTemplates":
                    {
                        break;
                    }

                case "RightsManagementTemplate":
                    {
                        break;
                    }

                case "RightsManagementLicense":
                    {
                        break;
                    }

                case "EditAllowed":
                    {
                        break;
                    }

                case "ReplyAllowed":
                    {
                        break;
                    }

                case "ReplyAllAllowed":
                    {
                        break;
                    }

                case "ForwardAllowed":
                    {
                        break;
                    }

                case "ModifyRecipientsAllowed":
                    {
                        break;
                    }

                case "ExtractAllowed":
                    {
                        break;
                    }

                case "PrintAllowed":
                    {
                        break;
                    }

                case "ExportAllowed":
                    {
                        break;
                    }

                case "ProgrammaticAccessAllowed":
                    {
                        break;
                    }

                case "RMOwner":
                    {
                        break;
                    }

                case "ContentExpiryDate":
                    {
                        break;
                    }

                case "TemplateID":
                    {
                        break;
                    }

                case "TemplateName":
                    {
                        break;
                    }

                case "TemplateDescription":
                    {
                        break;
                    }

                case "ContentOwner":
                    {
                        break;
                    }

                case "RemoveRightsManagementDistribution":
                    {
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
        #endregion
    }
}
#endregion