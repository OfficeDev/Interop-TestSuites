namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_ASRMAdapter
    {
        #region Verify transport
        /// <summary>
        /// This method is used to verify transport related requirement.
        /// </summary>
        private void VerifyTransport()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R2");

            // Verify MS-ASRM requirement: MS-ASRM_R2
            // ActiveSyncClient encodes XML request into WBXML and decodes WBXML to XML response, capture it directly if server responses succeed.
            Site.CaptureRequirement(
                2,
                @"[In Transport] The XML markup that constitutes the request body or the response body is transmitted between client and server by using Wireless Application Protocol (WAP) Binary XML (WBXML), as specified in [MS-ASWBXML].");
        }
        #endregion

        #region Verify requirements about RightsManagementLicense
        /// <summary>
        /// Verify requirements about RightsManagementLicense.
        /// </summary>
        /// <param name="rightsManagementLicense">The RightsManagementLicense element.</param>
        private void VerifyRightsManagementLicense(Response.RightsManagementLicense rightsManagementLicense)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R108");

            // Verify MS-ASRM requirement: MS-ASRM_R108
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                108,
                @"[In RightsManagementLicense] The value of this element[RightsManagementLicense] is a container ([MS-ASDTYPE] section 2.2).");

            this.VerifyContainer();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R21");

            // Verify MS-ASRM requirement: MS-ASRM_R21
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                21,
                @"[In ContentExpiryDate] The ContentExpiryDate element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R284");

            // Verify MS-ASRM requirement: MS-ASRM_R284
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                284,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ContentExpiryDate (section 2.2.2.1). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R27");

            // Verify MS-ASRM requirement: MS-ASRM_R27
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                27,
                @"[In ContentExpiryDate] The value of this element[ContentExpiryDate] is a dateTime ([MS-ASDTYPE] section 2.3).");

            this.VerifyDateTime();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R28");

            // Verify MS-ASRM requirement: MS-ASRM_R28
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                28,
                @"[In ContentExpiryDate] The ContentExpiryDate element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R29");

            // Verify MS-ASRM requirement: MS-ASRM_R29
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                29,
                @"[In ContentOwner] The ContentOwner element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R285");

            // Verify MS-ASRM requirement: MS-ASRM_R285
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                285,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ContentOwner (section 2.2.2.2). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R32");

            // Verify MS-ASRM requirement: MS-ASRM_R32
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                32,
                @"[In ContentOwner] The value of this element[ContentOwner] is a NonEmptyStringType, as specified in section 2.2.");

            this.VerifyNonEmptyString(rightsManagementLicense.ContentOwner);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R33");

            // Verify MS-ASRM requirement: MS-ASRM_R33
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                33,
                @"[In ContentOwner] The ContentOwner element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R275, the actual length of ContentOwner element: {0}", rightsManagementLicense.ContentOwner.Length);

            // Verify MS-ASRM requirement: MS-ASRM_R275
            Site.CaptureRequirementIfIsTrue(
                rightsManagementLicense.ContentOwner.Length < 320,
                275,
                @"[In ContentOwner] The length of the ContentOwner value is less than 320 characters.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R35");

            // Verify MS-ASRM requirement: MS-ASRM_R35
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                35,
                @"[In EditAllowed] The EditAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R286");

            // Verify MS-ASRM requirement: MS-ASRM_R286
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                286,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]EditAllowed (section 2.2.2.3). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R37");

            // Verify MS-ASRM requirement: MS-ASRM_R37
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                37,
                @"[In EditAllowed] The value of this element[EditAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            this.VerifyBoolean();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R43");

            // Verify MS-ASRM requirement: MS-ASRM_R43
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                43,
                @"[In EditAllowed] The EditAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R45");

            // Verify MS-ASRM requirement: MS-ASRM_R45
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                45,
                @"[In ExportAllowed] The ExportAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R287");

            // Verify MS-ASRM requirement: MS-ASRM_R287
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                287,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ExportAllowed (section 2.2.2.4). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R47");

            // Verify MS-ASRM requirement: MS-ASRM_R47
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                47,
                @"[In ExportAllowed] The value of this element[ExportAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R53");

            // Verify MS-ASRM requirement: MS-ASRM_R53
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                53,
                @"[In ExportAllowed] The ExportAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R54");

            // Verify MS-ASRM requirement: MS-ASRM_R54
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                54,
                @"[In ExtractAllowed] The ExtractAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R288");

            // Verify MS-ASRM requirement: MS-ASRM_R288
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                288,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ExtractAllowed (section 2.2.2.5). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R56");

            // Verify MS-ASRM requirement: MS-ASRM_R56
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                56,
                @"[In ExtractAllowed] The value of this element[ExtractAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R58");

            // Verify MS-ASRM requirement: MS-ASRM_R58
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                58,
                @"[In ExtractAllowed] The ExtractAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R59");

            // Verify MS-ASRM requirement: MS-ASRM_R59
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                59,
                @"[In ForwardAllowed] The ForwardAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R289");

            // Verify MS-ASRM requirement: MS-ASRM_R289
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                289,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ForwardAllowed (section 2.2.2.6). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R61");

            // Verify MS-ASRM requirement: MS-ASRM_R61
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                61,
                @"[In ForwardAllowed] The value of this element[ForwardAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R63");

            // Verify MS-ASRM requirement: MS-ASRM_R63
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                63,
                @"[In ForwardAllowed] The ForwardAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R64");

            // Verify MS-ASRM requirement: MS-ASRM_R64
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                64,
                @"[In ModifyRecipientsAllowed] The ModifyRecipientsAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R290");

            // Verify MS-ASRM requirement: MS-ASRM_R290
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                290,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]ModifyRecipientsAllowed (section 2.2.2.7). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R66");

            // Verify MS-ASRM requirement: MS-ASRM_R66
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                66,
                @"[In ModifyRecipientsAllowed] The value of this element[ModifyRecipientsAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R68");

            // Verify MS-ASRM requirement: MS-ASRM_R68
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                68,
                @"[In ModifyRecipientsAllowed] The ModifyRecipientsAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R69");

            // Verify MS-ASRM requirement: MS-ASRM_R69
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                69,
                @"[In Owner] The Owner element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R291");

            // Verify MS-ASRM requirement: MS-ASRM_R291
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                291,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]Owner (section 2.2.2.8). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R71");

            // Verify MS-ASRM requirement: MS-ASRM_R71
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                71,
                @"[In Owner] The value of this element[Owner] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R76");

            // Verify MS-ASRM requirement: MS-ASRM_R76
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                76,
                @"[In Owner] The Owner element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R77");

            // Verify MS-ASRM requirement: MS-ASRM_R77
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                77,
                @"[In PrintAllowed] The PrintAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R292");

            // Verify MS-ASRM requirement: MS-ASRM_R292
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                292,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]PrintAllowed (section 2.2.2.9). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R80");

            // Verify MS-ASRM requirement: MS-ASRM_R80
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                80,
                @"[In PrintAllowed] The value of this element[PrintAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R82");

            // Verify MS-ASRM requirement: MS-ASRM_R82
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                82,
                @"[In PrintAllowed] The PrintAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R83");

            // Verify MS-ASRM requirement: MS-ASRM_R83
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                83,
                @"[In ProgrammaticAccessAllowed] The ProgrammaticAccessAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R293");

            // Verify MS-ASRM requirement: MS-ASRM_R293
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                293,
                @"[In RightsManagementLicense][The RightsManagementLicense element can only have the following child elements:]ProgrammaticAccessAllowed (section 2.2.2.10). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R85");

            // Verify MS-ASRM requirement: MS-ASRM_R85
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                85,
                @"[In ProgrammaticAccessAllowed] The value of this element[ProgrammaticAccessAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R89");

            // Verify MS-ASRM requirement: MS-ASRM_R89
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                89,
                @"[In ProgrammaticAccessAllowed] The ProgrammaticAccessAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R96");

            // Verify MS-ASRM requirement: MS-ASRM_R96
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                96,
                @"[In ReplyAllAllowed] The ReplyAllAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R294");

            // Verify MS-ASRM requirement: MS-ASRM_R294
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                294,
                @"[In RightsManagementLicense][The RightsManagementLicense element can only have the following child elements:]ReplyAllAllowed (section 2.2.2.12). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R98");

            // Verify MS-ASRM requirement: MS-ASRM_R98
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                98,
                @"[In ReplyAllAllowed] The value of this element[ReplyAllAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R100");

            // Verify MS-ASRM requirement: MS-ASRM_R100
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                100,
                @"[In ReplyAllAllowed] The ReplyAllAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R101");

            // Verify MS-ASRM requirement: MS-ASRM_R101
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                101,
                @"[In ReplyAllowed] The ReplyAllowed element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R295");

            // Verify MS-ASRM requirement: MS-ASRM_R295
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                295,
                @"[In RightsManagementLicense][The RightsManagementLicense element can only have the following child elements:]ReplyAllowed (section 2.2.2.13). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R103");

            // Verify MS-ASRM requirement: MS-ASRM_R103
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                103,
                @"[In ReplyAllowed] The value of this element[ReplyAllowed] is a boolean ([MS-ASDTYPE] section 2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R105");

            // Verify MS-ASRM requirement: MS-ASRM_R105
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                105,
                @"[In ReplyAllowed] The ReplyAllowed element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R143");

            // Verify MS-ASRM requirement: MS-ASRM_R143
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                143,
                @"[In TemplateDescription (RightsManagementLicense)] The TemplateDescription element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R296");

            // Verify MS-ASRM requirement: MS-ASRM_R296
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                296,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]TemplateDescription (section 2.2.2.18.1). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R142");

            // Verify MS-ASRM requirement: MS-ASRM_R142
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                142,
                @"[In TemplateDescription] The value of this element[TemplateDescription] is a NonEmptyStringType, as specified in section 2.2.");

            this.VerifyNonEmptyString(rightsManagementLicense.TemplateDescription);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R146");

            // Verify MS-ASRM requirement: MS-ASRM_R146
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                146,
                @"[In TemplateDescription (RightsManagementLicense)] The TemplateDescription element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R277, the actual length of TemplateDescription (RightsManagementLicense) element: {0}", rightsManagementLicense.TemplateDescription.Length);

            // Verify MS-ASRM requirement: MS-ASRM_R277
            Site.CaptureRequirementIfIsTrue(
                rightsManagementLicense.TemplateDescription.Length < 10240,
                277,
                @"[In TemplateDescription (RightsManagementLicense)] The length of the TemplateDescription[RightsManagementLicense] element is less than 10240 characters.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R307");

            // Verify MS-ASRM requirement: MS-ASRM_R307
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                307,
                @"[In TemplateID] The value of this element[TemplateID] is a NonEmptyStringType, as specified in section 2.2.");

            this.VerifyNonEmptyString(rightsManagementLicense.TemplateID);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R159");

            // Verify MS-ASRM requirement: MS-ASRM_R159
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                159,
                @"[In TemplateID (RightsManagementLicense)] The TemplateID element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R297");

            // Verify MS-ASRM requirement: MS-ASRM_R297
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                297,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]TemplateID (section 2.2.2.19.1). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R160");

            // Verify MS-ASRM requirement: MS-ASRM_R160
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                160,
                @"[In TemplateID (RightsManagementLicense)] It[TemplateID] contains a string that identifies the rights policy template represented by the parent RightsManagementLicense element.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R161");

            // Verify MS-ASRM requirement: MS-ASRM_R161
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                161,
                @"[In TemplateID (RightsManagementLicense)] The TemplateID element[RightsManagementLicense] has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R172");

            // Verify MS-ASRM requirement: MS-ASRM_R172
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                172,
                @"[In TemplateName] The value of this element[TemplateName] is a NonEmptyStringType, as specified in section 2.2.");

            this.VerifyNonEmptyString(rightsManagementLicense.TemplateName);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R173");

            // Verify MS-ASRM requirement: MS-ASRM_R173
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                173,
                @"[In TemplateName (RightsManagementLicense)] The TemplateName element is a required child element of the RightsManagementLicense element (section 2.2.2.14).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R298");

            // Verify MS-ASRM requirement: MS-ASRM_R298
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                298,
                @"[In RightsManagementLicense] [The RightsManagementLicense element can only have the following child elements:]TemplateName (section 2.2.2.20.1). This element is required.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R176");

            // Verify MS-ASRM requirement: MS-ASRM_R176
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                176,
                @"[In TemplateName (RightsManagementLicense)] The TemplateName[RightsManagementLicense] element has no child elements.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R280, the actual length of TemplateName (RightsManagementLicense) element: {0}", rightsManagementLicense.TemplateName.Length);

            // Verify MS-ASRM requirement: MS-ASRM_R280
            Site.CaptureRequirementIfIsTrue(
                rightsManagementLicense.TemplateName.Length < 256,
                280,
                @"[In TemplateName (RightsManagementLicense)] The length of the TemplateName[RightsManagementLicense] element is less than 256 characters.");
        }
        #endregion

        #region Verify requirements about Settings response
        /// <summary>
        /// Verify the rights-managed requirements about Settings response.
        /// </summary>
        /// <param name="settingsResponse">The response of Settings command.</param>
        private void VerifySettingsResponse(SettingsResponse settingsResponse)
        {
            // Verify the schema of MS-ASRM.
            if (settingsResponse.ResponseData.RightsManagementInformation != null)
            {
                if (settingsResponse.ResponseData.RightsManagementInformation.Get != null)
                {
                    this.VerifyRightsManagementTemplates(settingsResponse.ResponseData.RightsManagementInformation.Get);
                }
            }
        }

        /// <summary>
        /// Verify requirements about RightsManagementTemplates
        /// </summary>
        /// <param name="rightsManagementTemplatesResponse">The context of rights management templates</param>
        private void VerifyRightsManagementTemplates(Response.SettingsRightsManagementInformationGet rightsManagementTemplatesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R137");

            // Verify MS-ASRM requirement: MS-ASRM_R137
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                137,
                @"[In RightsManagementTemplates] The value of this element[RightsManagementTemplates] is a container ([MS-ASDTYPE] section 2.2).");

            this.VerifyContainer();

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R302");

            // Verify MS-ASRM requirement: MS-ASRM_R302
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                302,
                @"[In RightsManagementTemplates] The RightsManagementTemplates element can only have the following child element: RightsManagementTemplate (section 2.2.2.16). ");

            int lengthOfArray = rightsManagementTemplatesResponse.RightsManagementTemplates.Length;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R305, the actual number of RightsManagementTemplates element: {0}", lengthOfArray);
            
            // Verify MS-ASRM requirement: MS-ASRM_R305
            Site.CaptureRequirementIfIsTrue(
                lengthOfArray < 20,
                305,
                @"[In RightsManagementTemplates] The RightsManagementTemplate elements returned to the client is less than 20.");

            for (int i = 0; i < lengthOfArray; i++)
            {
                this.VerifyRightsManagementTemplate(rightsManagementTemplatesResponse.RightsManagementTemplates[i]);
            }
        }
        #endregion

        #region Verify requirements about Sync response
        /// <summary>
        /// Verify the rights-managed requirements about Sync response.
        /// </summary>
        /// <param name="sync">The wrapper class of Sync response.</param>
        private void VerifySyncResponse(Sync sync)
        {
            if (sync != null)
            {
                Site.Assert.IsNotNull(sync.Email, "The expected rights-managed e-mail message should not be null.");
                if (sync.Email.RightsManagementLicense != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R208");

                    // Verify MS-ASRM requirement: MS-ASRM_R208
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        208,
                        @"[In Sending Rights-Managed E-Mail Messages to the Client] To respond to a Sync command request message that includes the RightsManagementSupport element, the server includes the RightsManagementLicense element and its child elements in the Sync command response message.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R209");

                    // Verify MS-ASRM requirement: MS-ASRM_R209
                    Site.CaptureRequirementIfIsTrue(
                        this.activeSyncClient.ValidationResult,
                        209,
                        @"[In Sending Rights-Managed E-Mail Messages to the Client] In a Sync command response, the RightsManagementLicense element is included as a child of the sync:ApplicationData element ([MS-ASCMD] section 2.2.3.11).");

                    this.VerifyRightsManagementLicense(sync.Email.RightsManagementLicense);
                }
            }
        }
        #endregion

        #region Verify requirements about ItemOperations response
        /// <summary>
        /// Verify the rights-managed requirements about ItemOperations response.
        /// </summary>
        /// <param name="itemOperationsStore">The wrapper class for the fetched result of ItemOperations command.</param>
        private void VerifyItemOperationsResponse(ItemOperationsStore itemOperationsStore)
        {
            foreach (ItemOperations itemOperations in itemOperationsStore.Items)
            {
                if (itemOperations.Email != null)
                {
                    if (itemOperations.Email.RightsManagementLicense != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R210");

                        // Verify MS-ASRM requirement: MS-ASRM_R210
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            210,
                            @"[In Sending Rights-Managed E-Mail Messages to the Client] In an ItemOperations command response, the RightsManagementLicense element is included as a child of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.128.1).");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R410");

                        // Verify MS-ASRM requirement: MS-ASRM_R410
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            410,
                            @"[In Sending Rights-Managed E-Mail Messages to the Client] To respond to an ItemOperations command request message that includes the RightsManagementSupport element, the server includes the RightsManagementLicense element and its child elements in the ItemOperations command response message.");

                        this.VerifyRightsManagementLicense(itemOperations.Email.RightsManagementLicense);
                    }
                }
            }
        }
        #endregion

        #region Verify requirements about Search response
        /// <summary>
        /// Verify the rights-managed requirements about Search response.
        /// </summary>
        /// <param name="searchStore">The wrapper class for the result of Search command.</param>
        private void VerifySearchResponse(SearchStore searchStore)
        {
            foreach (Search search in searchStore.Results)
            {
                if (search.Email != null)
                {
                    if (search.Email.RightsManagementLicense != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R211");

                        // Verify MS-ASRM requirement: MS-ASRM_R211
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            211,
                            @"[In Sending Rights-Managed E-Mail Messages to the Client] In a Search command response, the RightsManagementLicense element is included as a child of the search:Properties element ([MS-ASCMD] section 2.2.3.128.2).");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R411");

                        // Verify MS-ASRM requirement: MS-ASRM_R411
                        Site.CaptureRequirementIfIsTrue(
                            this.activeSyncClient.ValidationResult,
                            411,
                            @"[In Sending Rights-Managed E-Mail Messages to the Client] To respond to a Search command request message that includes the RightsManagementSupport element, the server includes the RightsManagementLicense element and its child elements in the Search command response message.");

                        this.VerifyRightsManagementLicense(search.Email.RightsManagementLicense);
                    }
                }
            }
        }
        #endregion

        #region Verify requirements about RightsManagementTemplate
        /// <summary>
        /// Verify requirements about RightsManagementTemplate included in RightsManagementTemplates.
        /// </summary>
        /// <param name="rightsManagementTemplate">The context of rights management template in RightsManagementTemplates element.</param>
        private void VerifyRightsManagementTemplate(Response.RightsManagementTemplatesRightsManagementTemplate rightsManagementTemplate)
        {
            if (rightsManagementTemplate != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R132");

                // Verify MS-ASRM requirement: MS-ASRM_R132
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    132,
                    @"[In RightsManagementTemplate] The value of this element[RightsManagementTemplate] is a container ([MS-ASDTYPE] section 2.2).");

                this.VerifyContainer();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R148");

                // Verify MS-ASRM requirement: MS-ASRM_R148
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    148,
                    @"[In TemplateDescription (RightsManagementTemplate)] The TemplateDescription element is a required child element of the RightsManagementTemplate element (section 2.2.2.16).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R133");

                // Verify MS-ASRM requirement: MS-ASRM_R133
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    133,
                    @"[In RightsManagementTemplate] [The RightsManagementTemplate element can have only one of each of the following child elements:]	TemplateDescription (section 2.2.2.18.2). This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R151");

                // Verify MS-ASRM requirement: MS-ASRM_R151
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    151,
                    @"[In TemplateDescription (RightsManagementTemplate)] The TemplateDescription element has no child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R278, the actual length of TemplateDescription element: {0}", rightsManagementTemplate.TemplateDescription.Length);

                // Verify MS-ASRM requirement: MS-ASRM_R278
                Site.CaptureRequirementIfIsTrue(
                    rightsManagementTemplate.TemplateDescription.Length < 10240,
                    278,
                    @"[In TemplateDescription (RightsManagementTemplate)] The length of the TemplateDescription[RightsManagementTemplate] element is less than 10240 characters.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R162");

                // Verify MS-ASRM requirement: MS-ASRM_R162
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    162,
                    @"[In TemplateID (RightsManagementTemplate)] The TemplateID element is a required child element of the RightsManagementTemplate element (section 2.2.2.16).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R163");

                // Verify MS-ASRM requirement: MS-ASRM_R163
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    163,
                    @"[In TemplateID (RightsManagementTemplate)] It[TemplateID] contains a string that identifies the rights policy template represented by the parent RightsManagementTempalte element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R164");

                // Verify MS-ASRM requirement: MS-ASRM_R164
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    164,
                    @"[In TemplateID (RightsManagementTemplate)] The TemplateID[RightsManagementTemplate] element has no child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R134");

                // Verify MS-ASRM requirement: MS-ASRM_R134
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    134,
                    @"[In RightsManagementTemplate] [The RightsManagementTemplate element can have only one of each of the following child elements:]TemplateID (section 2.2.2.19.2). This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R178");

                // Verify MS-ASRM requirement: MS-ASRM_R178
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    178,
                    @"[In TemplateName (RightsManagementTemplate)] The TemplateName element is a required child element of the RightsManagementTemplate element (section 2.2.2.16).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R181");

                // Verify MS-ASRM requirement: MS-ASRM_R181
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    181,
                    @"[In TemplateName (RightsManagementTemplate)] The TemplateName[RightsManagementTemplate] element has no child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R135");

                // Verify MS-ASRM requirement: MS-ASRM_R135
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    135,
                    @"[In RightsManagementTemplate] [The RightsManagementTemplate element can have only one of each of the following child elements:]	TemplateName (section 2.2.2.20.2). This element is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R279, the actual length of TemplateName element: {0}", rightsManagementTemplate.TemplateName.Length);

                // Verify MS-ASRM requirement: MS-ASRM_R279
                Site.CaptureRequirementIfIsTrue(
                    rightsManagementTemplate.TemplateName.Length < 256,
                    279,
                    @"[In TemplateName (RightsManagementTemplate)] The length of the TemplateName[RightsManagementTemplate] element is less than 256 characters.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASRM_R317");

                // Verify MS-ASRM requirement: MS-ASRM_R317
                Site.CaptureRequirementIfIsTrue(
                    this.activeSyncClient.ValidationResult,
                    317,
                    @"[In RightsManagementTemplate] The RightsManagementTemplate element can have only one of each of the following child elements[TemplateDescription, TemplateID, TemplateName]");
            }
        }
        #endregion

        #region Verify requirements in MS-ASDTYPE
        /// <summary>
        /// This method is used to verify the boolean related requirements.
        /// </summary>
        private void VerifyBoolean()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R4");

            // If the validation is successful, then MS-ASDTYPE_R4 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                4,
                @"[In boolean Data Type] It [a boolean] is declared as an element with a type attribute of ""boolean"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R5");

            // When the boolean type element successfully passed schema validation, 
            // it is proved that the lower layer EAS client stack library did the WBXML decoding successfully,then MS-ASDTYPE_R5 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
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
        /// This method is used to verify the container related requirements.
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
        /// This method is used to verify the datetime related requirements.
        /// </summary>
        private void VerifyDateTime()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R12");

            // If the validation is successful, then MS-ASDTYPE_R12 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.activeSyncClient.ValidationResult,
                "MS-ASDTYPE",
                12,
                @"[In dateTime Data Type] It [dateTime]is declared as an element whose type attribute is set to ""dateTime"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R15");

            // When the dateTime type element successfully passed schema validation, 
            // it is proved that the lower layer EAS client stack library did the WBXML decoding successfully, then MS-ASDTYPE_R15 can be captured.
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

            // If MS-ASDTYPE_R15 can be captured successfully, then MS-ASDTYPE_R16 can be captured directly.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R16");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R16
            Site.CaptureRequirement(
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
        /// This method is used to verify whether the string is a non-empty string.
        /// </summary>
        /// <param name="value">The string value</param>
        private void VerifyNonEmptyString(string value)
        {
            // Check the string length
            Site.Assert.IsTrue(value.Length >= 1, "The length of the string should be at least equal or greater than 1, actual length: {0}.", value.Length);
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
                    Site.Assert.IsTrue(codepage >= 0 && codepage <= 24, "Code page value should between 0-24, the actual value is :{0}", codepage);

                    // Capture the requirements in RightsManagement namespace
                    if (24 == codepage)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R34");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R34
                        Site.CaptureRequirementIfAreEqual<string>(
                            "rightsmanagement",
                            codePageName.ToLower(CultureInfo.CurrentCulture),
                            "MS-ASWBXML",
                            34,
                            @"[In Code Pages] [This algorithm supports] [Code page] 24[that indicates] [XML namespace] RightsManagement");

                        this.VerifyRequirementsRelateToCodePage24(codepage, tagName, token);
                    }
                }
            }
        }

        /// <summary>
        /// Verify the tags and tokens in WBXML code page 24.
        /// </summary>
        /// <param name="codePageNumber">The code page number.</param>
        /// <param name="tagName">The tag name that needs to be verified.</param>
        /// <param name="token">The token that needs to be verified.</param>
        private void VerifyRequirementsRelateToCodePage24(int codePageNumber, string tagName, byte token)
        {
            switch (tagName)
            {
                case "RightsManagementTemplates":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R633");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R633
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x06,
                            token,
                            "MS-ASWBXML",
                            633,
                            @"[In Code Page 24: RightsManagement] [Tag name] RightsManagementTemplates [Token] 0x06 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "RightsManagementTemplate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R634");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R634
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x07,
                            token,
                            "MS-ASWBXML",
                            634,
                            @"[In Code Page 24: RightsManagement] [Tag name] RightsManagementTemplate [Token] 0x07 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "RightsManagementLicense":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R635");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R635
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x08,
                            token,
                            "MS-ASWBXML",
                            635,
                            @"[In Code Page 24: RightsManagement] [Tag name] RightsManagementLicense [Token] 0x08 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "EditAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R636");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R636
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x09,
                            token,
                            "MS-ASWBXML",
                            636,
                            @"[In Code Page 24: RightsManagement] [Tag name] EditAllowed [Token] 0x09 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ReplyAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R637");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R637
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0A,
                            token,
                            "MS-ASWBXML",
                            637,
                            @"[In Code Page 24: RightsManagement] [Tag name] ReplyAllowed [Token] 0x0A [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ReplyAllAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R638");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R638
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0B,
                            token,
                            "MS-ASWBXML",
                            638,
                            @"[In Code Page 24: RightsManagement] [Tag name] ReplyAllAllowed [Token] 0x0B [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ForwardAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R639");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R639
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0C,
                            token,
                            "MS-ASWBXML",
                            639,
                            @"[In Code Page 24: RightsManagement] [Tag name] ForwardAllowed [Token] 0x0C [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ModifyRecipientsAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R640");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R640
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0D,
                            token,
                            "MS-ASWBXML",
                            640,
                            @"[In Code Page 24: RightsManagement] [Tag name] ModifyRecipientsAllowed [Token] 0x0D [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ExtractAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R641");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R641
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0E,
                            token,
                            "MS-ASWBXML",
                            641,
                            @"[In Code Page 24: RightsManagement] [Tag name] ExtractAllowed [Token] 0x0E [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "PrintAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R642");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R642
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x0F,
                            token,
                            "MS-ASWBXML",
                            642,
                            @"[In Code Page 24: RightsManagement] [Tag name] PrintAllowed [Token] 0x0F [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ExportAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R643");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R643
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x10,
                            token,
                            "MS-ASWBXML",
                            643,
                            @"[In Code Page 24: RightsManagement] [Tag name] ExportAllowed [Token] 0x10 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ProgrammaticAccessAllowed":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R644");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R644
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x11,
                            token,
                            "MS-ASWBXML",
                            644,
                            @"[In Code Page 24: RightsManagement] [Tag name] ProgrammaticAccessAllowed [Token] 0x11 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "Owner":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R645");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R645
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x12,
                            token,
                            "MS-ASWBXML",
                            645,
                            @"[In Code Page 24: RightsManagement] [Tag name] Owner [Token] 0x12 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ContentExpiryDate":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R646");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R646
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x13,
                            token,
                            "MS-ASWBXML",
                            646,
                            @"[In Code Page 24: RightsManagement] [Tag name] ContentExpiryDate [Token] 0x13 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "TemplateID":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R647");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R647
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x14,
                            token,
                            "MS-ASWBXML",
                            647,
                            @"[In Code Page 24: RightsManagement] [Tag name] TemplateID [Token] 0x14 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "TemplateName":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R648");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R648
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x15,
                            token,
                            "MS-ASWBXML",
                            648,
                            @"[In Code Page 24: RightsManagement] [Tag name] TemplateName [Token] 0x15 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "TemplateDescription":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R649");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R649
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x16,
                            token,
                            "MS-ASWBXML",
                            649,
                            @"[In Code Page 24: RightsManagement] [Tag name] TemplateDescription [Token] 0x16 [supports protocol versions] 14.1, 16.0");

                        break;
                    }

                case "ContentOwner":
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R650");

                        // Verify MS-ASWBXML requirement: MS-ASWBXML_R650
                        Site.CaptureRequirementIfAreEqual<byte>(
                            0x17,
                            token,
                            "MS-ASWBXML",
                            650,
                            @"[In Code Page 24: RightsManagement] [Tag name] ContentOwner [Token] 0x17 [supports protocol versions] 14.1, 16.0");

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