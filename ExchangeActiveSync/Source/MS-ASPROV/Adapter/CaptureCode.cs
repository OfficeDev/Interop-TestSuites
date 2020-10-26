namespace Microsoft.Protocols.TestSuites.MS_ASPROV
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-ASPROV.
    /// </summary>
    public partial class MS_ASPROVAdapter
    {
        /// <summary>
        /// Verify Provision command requirements.
        /// </summary>
        /// <param name="provisionResponse">Provision response</param>
        private void VerifyProvisionCommandRequirements(ProvisionResponse provisionResponse)
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");
                        
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R60010");

            // Verify MS-ASPROV requirement: MS-ASPROV_R60010
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                60010,
                @"[In Transport] The encoded XML block containing the command and parameter elements is transmitted[ in either the request body of a request, or] in the response body of a response.");

            // Verify requirements of AccountOnlyRemoteWipe.
            if (null != provisionResponse.ResponseData.AccountOnlyRemoteWipe)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R6901");

                // Verify MS-ASPROV requirement: MS-ASPROV_R6901
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    6901,
                    @"[In AccountOnlyRemoteWipe] The AccountOnlyRemoteWipe element is an optional container ([MS-ASDTYPE] section 2.2) element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R6903");

                // Verify MS-ASPROV requirement: MS-ASPROV_R6903
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    6903,
                    @"[In AccountOnlyRemoteWipe] A server response MUST NOT include any child elements in the AccountOnlyRemoteWipe element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R66605");

                // Verify MS-ASPROV requirement: MS-ASPROV_R66605
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    66605,
                    @"[In Responding to an Account Only Remote Wipe Directive Acknowledgement] The server's response is in the following format.
<Provision>
	<Status>...</Status>
	<AccountOnlyRemoteWipe/>
</Provision>");

                this.VerifyContainerStructure();
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R319");

            // Verify MS-ASPROV requirement: MS-ASPROV_R319
            // The ResponseData is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData,
                319,
                @"[In Provision] The Provision element is a required container ([MS-ASDTYPE] section 2.2) element in a provisioning request and response that specifies the capabilities and permissions of a device.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R320");

            // Verify MS-ASPROV requirement: MS-ASPROV_R320
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                320,
                @"[In Provision] The Provision element has the following child elements:
settings:DeviceInformation (section 2.2.2.53)
Status (section 2.2.2.54.2)
Policies (section 2.2.2.40)
RemoteWipe (section 2.2.2.45)
AccountOnlyRemoteWipe (section 2.2.2.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R759");

            // Verify MS-ASPROV requirement: MS-ASPROV_R759
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                759,
                @"[In Status] The Status element is a child element of the Provision element (section 2.2.2.44).");

            // Verify requirements of DeviceInformation.
            if (null != provisionResponse.ResponseData.DeviceInformation)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R366");

                // Verify MS-ASPROV requirement: MS-ASPROV_R366
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    366,
                    @"[In settings:DeviceInformation] It [settings:DeviceInformation element] is a child of the Provision element (section 2.2.2.44).");

                this.VerifyContainerStructure();
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R390");

            // Verify MS-ASPROV requirement: MS-ASPROV_R390
            // The Status element in Provision is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                provisionResponse.ResponseData.Status,
                390,
                @"[In Status (Provision)] The Status (Provision) element is a required child element of the Provision element in command responses.");

            this.VerifyUnsignedByteStructure(provisionResponse.ResponseData.Status);

            if (provisionResponse.ResponseData.Status < 100)
            {
                bool isVerifiedR393 = provisionResponse.ResponseData.Status == 1 || provisionResponse.ResponseData.Status == 2 || provisionResponse.ResponseData.Status == 3;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR393,
                    393,
                    @"[In Status (Provision)] The following table lists valid values [1,2,3] for the Status (Provision) element when it is the child of the Provision element.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R391");

            // Verify MS-ASPROV requirement: MS-ASPROV_R391
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                391,
                @"[In Status (Provision)] The value of this element [Status (Provision)] is an unsignedByte ([MS-ASDTYPE] section 2.8).");

            // Verify requirements of Policies.
            if (null != provisionResponse.ResponseData.Policies)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R303");

                // Verify MS-ASPROV requirement: MS-ASPROV_R303
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    303,
                    @"[In Policies] The Policies element is a required container ([MS-ASDTYPE] section 2.2) element that specifies a collection of security policies.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R304");

                // Verify MS-ASPROV requirement: MS-ASPROV_R304
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    304,
                    @"[In Policies] It [Policies element] is a child of the Provision element (section 2.2.2.44).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R307");

                // Verify MS-ASPROV requirement: MS-ASPROV_R307
                // The Policy element is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    provisionResponse.ResponseData.Policies.Policy,
                    307,
                    @"[In Policy] The Policy element is a required container ([MS-ASDTYPE] section 2.2) element that specifies a policy.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R306");

                // Verify MS-ASPROV requirement: MS-ASPROV_R306
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    306,
                    @"[In Policies] The Policies element has only the following child element: Policy (section 2.2.2.41): At least one element of this type is required.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R308");

                // Verify MS-ASPROV requirement: MS-ASPROV_R308
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    308,
                    @"[In Policy] It [Policy element] is a child of the Policies element (section 2.2.2.40).");

                // Verify requirements of PolicyType.
                if (null != provisionResponse.ResponseData.Policies.Policy.PolicyType)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R317");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R317
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        317,
                        @"[In PolicyType] The PolicyType element is a child element of type string ([MS-ASDTYPE] section 2.7) of the Policy element (section 2.2.2.41).");

                    if (provisionResponse.ResponseData.Policies.Policy.Status != 3)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R318");

                        // Verify MS-ASPROV requirement: MS-ASPROV_R318
                        bool isR318Satisfied = provisionResponse.ResponseData.Policies.Policy.PolicyType.Equals("MS-WAP-Provisioning-XML") || provisionResponse.ResponseData.Policies.Policy.PolicyType.Equals("MS-EAS-Provisioning-WBXML");

                        Site.CaptureRequirementIfIsTrue(
                            isR318Satisfied,
                            318,
                            @"[In PolicyType] The value of the PolicyType element MUST be one of the values specified in the following table.
[MS-WAP-Provisioning-XML
MS-EAS-Provisioning-WBXML]");
                    }

                    this.VerifyStringStructure();
                }

                // Verify requirements of PolicyKey.
                if (null != provisionResponse.ResponseData.Policies.Policy.PolicyKey)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R313");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R313
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        313,
                        @"[In PolicyKey] It [PolicyKey] is a child element of the Policy element (section 2.2.2.41).");

                    this.VerifyStringStructure();

                    if (Common.IsRequirementEnabled(711, Site))
                    {
                        uint uintPolicyKey;

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R711");

                        // Verify MS-ASPROV requirement: MS-ASPROV_R711
                        // The PolicyKey could be parsed to unsigned integer, so this requirement can be captured.
                        Site.CaptureRequirementIfIsTrue(
                            uint.TryParse(provisionResponse.ResponseData.Policies.Policy.PolicyKey, out uintPolicyKey),
                            711,
                            @"[In Appendix B: Product Behavior] The value of the PolicyKey element is a string representation of a 32-bit unsigned integer. (Exchange 2007 and above follow this behavior.)");
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R374");

                // Verify MS-ASPROV requirement: MS-ASPROV_R374
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    374,
                    @"[In Status (Policy)] The Status element is a required child of the Policy element in command responses.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R370");

                // Verify MS-ASPROV requirement: MS-ASPROV_R370
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    370,
                    @"[In Status] The Status element is a child element of the Policy element (section 2.2.2.41).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R375");

                // Verify MS-ASPROV requirement: MS-ASPROV_R375
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    375,
                    @"[In Status (Policy)] In a command response, the value of this element [Status (Policy)] is an unsignedByte ([MS-ASDTYPE] section 2.8).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R377");

                Common.VerifyActualValues("Status(Policy)", new string[] { "1", "2", "3", "4", "5" }, provisionResponse.ResponseData.Policies.Policy.Status.ToString(), Site);

                // Verify MS-ASPROV requirement: MS-ASPROV_R377
                // The actual value of Status element is one of the valid values, so this requirement can be captured.
                Site.CaptureRequirement(
                    377,
                    @"[In Status (Policy)] The following table lists valid values [1,2,3,4,5] for the Status (Policy) element when it is the child of the Policy element in the response from the server to the client.");

                this.VerifyUnsignedByteStructure(provisionResponse.ResponseData.Policies.Policy.Status);

                // Verify requirements of Data.
                if (null != provisionResponse.ResponseData.Policies.Policy.Data)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R208");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R208
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        208,
                        @"[In Data (container Data Type)] The Data element as a container data type ([MS-ASDTYPE] section 2.2) contains a child element in which the policy settings for a device are specified. ");

                    this.Site.CaptureRequirementIfAreEqual<string>(
                        "MS-EAS-Provisioning-WBXML",
                        provisionResponse.ResponseData.Policies.Policy.PolicyType,
                        966,
                        @"[In PolicyType] Value MS-EAS-Provisioning-WBXML meaning The contents of the Data element are formatted according to the Exchange ActiveSync provisioning WBXML schema, as specified in section 2.2.2.24.1.");

                    if (Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("12.1") ||
                        Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.0") ||
                        Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("14.1") ||
                        Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.0") || 
                        Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site).Equals("16.1"))


                    {
                        this.Site.CaptureRequirementIfAreEqual<string>(
                            "MS-EAS-Provisioning-WBXML",
                            provisionResponse.ResponseData.Policies.Policy.PolicyType,
                            971,
                            @"[In PolicyType] The value ""MS-EAS-Provisioning-WBXML"" is used with protocol versions 12.0, 12.1, 14.0, 14.1, 16.0 and 16.1.");
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R210");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R210
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        210,
                        @"[In Data (container Data Type)] As a container data type, the Data element has only the following child element: EASProvisionDoc (section 2.2.2.28): One instance of this element is required.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R232");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R232
                    // The EASProvisionDoc element is not null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc,
                        232,
                        @"[In EASProvisionDoc] The EASProvisionDoc element is a required container ([MS-ASDTYPE] section 2.2) element that specifies the collection of security settings for device provisioning.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R233");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R233
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        233,
                        @"[In EASProvisionDoc] It [EASProvisionDoc element] is a child of the Data element (section 2.2.2.24.1).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R234");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R234
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        234,
                        @"[In EASProvisionDoc] A command response has a minimum of one EASProvisionDoc element per Data element.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R235");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R235
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        235,
                        @"[In EASProvisionDoc] The EASProvisionDoc element has only the following child elements:
AllowBluetooth (section 2.2.2.2)
AllowBrowser (section 2.2.2.3)
AllowCamera (section 2.2.2.4)
AllowConsumerEmail (section 2.2.2.5)
AllowDesktopSync (section 2.2.2.6)
AllowHTMLEmail (section 2.2.2.7)
AllowInternetSharing (section 2.2.2.8)
AllowIrDA (section 2.2.2.9)
AllowPOPIMAPEmail (section 2.2.2.10)
AllowRemoteDesktop (section 2.2.2.11)
AllowSimpleDevicePassword (section 2.2.2.12)
AllowSMIMEEncryptionAlgorithmNegotiation (section 2.2.2.13)
AllowSMIMESoftCerts (section 2.2.2.14)
AllowStorageCard (section 2.2.2.15)
AllowTextMessaging (section 2.2.2.16)
AllowUnsignedApplications (section 2.2.2.17)
AllowUnsignedInstallationPackages (section 2.2.2.18)
AllowWifi (section 2.2.2.19)
AlphanumericDevicePasswordRequired (section 2.2.2.20)
ApprovedApplicationList (section 2.2.2.22)
AttachmentsEnabled (section 2.2.2.23)
DevicePasswordEnabled (section 2.2.2.25)
DevicePasswordExpiration (section 2.2.2.26)
DevicePasswordHistory (section 2.2.2.27)
MaxAttachmentSize (section 2.2.2.30)
MaxCalendarAgeFilter (section 2.2.2.31)
MaxDevicePasswordFailedAttempts (section 2.2.2.32)
MaxEmailAgeFilter (section 2.2.2.33)
MaxEmailBodyTruncationSize (section 2.2.2.34)
MaxEmailHTMLBodyTruncationSize (section 2.2.2.35)
MaxInactivityTimeDeviceLock (section 2.2.2.36)
MinDevicePasswordComplexCharacters (section 2.2.2.37)
MinDevicePasswordLength (section 2.2.2.38)
PasswordRecoveryEnabled (section 2.2.2.39)
RequireDeviceEncryption (section 2.2.2.46)
RequireEncryptedSMIMEMessages (section 2.2.2.47)
RequireEncryptionSMIMEAlgorithm (section 2.2.2.48)
RequireManualSyncWhenRoaming (section 2.2.2.49)
RequireSignedSMIMEAlgorithm (section 2.2.2.50)
RequireSignedSMIMEMessages (section 2.2.2.51)
RequireStorageCardEncryption (section 2.2.2.52)
UnapprovedInROMApplicationList (section 2.2.2.55)");

                    this.VerifyEASProvisionDocElement(provisionResponse);

                    if (Common.IsRequirementEnabled(535, Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R535");

                        // Verify MS-ASPROV requirement: MS-ASPROV_R535
                        // The schema has been validated and RemoteWipe element is null, so this requirement can be captured.
                        Site.CaptureRequirementIfIsNull(
                            provisionResponse.ResponseData.RemoteWipe,
                            535,
                            @"[In Appendix B: Product Behavior] The implementation does respond to a security policy settings request in an initial Provision command request with a response in the following format. (Exchange 2007 and above follow this behavior.)
<Provision>
   <settings:DeviceInformation>
      <settings:Status>...</settings:Status>
   </settings:DeviceInformation>
   <Status>...</Status>
   <Policies>
      <Policy>
         <PolicyType>MS-EAS-Provisioning-WBXML</PolicyType>
         <Status>...</Status>
         <PolicyKey>...</PolicyKey>
         <Data>
            <EASProvisionDoc>
               ...
            </EASProvisionDoc>
         </Data>
      </Policy>
   </Policies>
</Provision>");
                    }

                    this.VerifyContainerStructure();
                }
                else
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R661");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R661
                    // The schema has been validated and the RemoteWipe element is null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNull(
                        provisionResponse.ResponseData.RemoteWipe,
                        661,
                        @"[In Responding to a Security Policy Settings Acknowledgment] If the level of compliance meets the server's requirements, the server response is in the following format.
<Provision>
   <Status>...</Status>
   <Policies>
      <Policy>
         <PolicyType>...</PolicyType>
         <Status>...</Status>
         <PolicyKey>...</PolicyKey>
      </Policy>
   </Policies>
</Provision>");
                }

                this.VerifyContainerStructure();
            }

            // Verify requirements of RemoteWipe.
            if (null != provisionResponse.ResponseData.RemoteWipe)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R321");

                // Verify MS-ASPROV requirement: MS-ASPROV_R321
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    321,
                    @"[In RemoteWipe] The RemoteWipe element is an optional container ([MS-ASDTYPE] section 2.2) element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R322");

                // Verify MS-ASPROV requirement: MS-ASPROV_R322
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    322,
                    @"[In RemoteWipe] A server response MUST NOT include any child elements in the RemoteWipe element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R760");

                // Verify MS-ASPROV requirement: MS-ASPROV_R760
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    760,
                    @"[In Status] The Status element is a child element of the RemoteWipe element (section 2.2.2.45).");

                if (Common.IsRequirementEnabled(758, Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R758");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R758
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        758,
                        @"[In Appendix B: Product Behavior] The implementation does respond to an empty initial Provision command request with a response in the following format. (Exchange 2007 and above follow this behavior.)
<Provision>
   <Status>...</Status>
   <RemoteWipe/>
</Provision>");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R666");

                // Verify MS-ASPROV requirement: MS-ASPROV_R666
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    666,
                    @"[In Responding to a Remote Wipe Directive Acknowledgment] The server's response is in the following format. <Provision> <Status>...</Status> <RemoteWipe/> </Provision>");

                this.VerifyContainerStructure();
            }

            this.VerifyContainerStructure();
        }

        #region Verify child elements of EASProvisionDoc element
        /// <summary>
        /// Verify child elements of EASProvisionDoc element.
        /// </summary>
        /// <param name="provisionResponse">Provision response</param>
        private void VerifyEASProvisionDocElement(ProvisionResponse provisionResponse)
        {
            // Get policy setting of provision command response.
            Dictionary<string, string> policiesSetting = AdapterHelper.GetPoliciesFromProvisionResponse(provisionResponse);

            if (policiesSetting.ContainsKey("AllowBluetooth"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R70");

                // Verify MS-ASPROV requirement: MS-ASPROV_R70
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    70,
                    @"[In AllowBluetooth] The AllowBluetooth element is an optional child element of type unsignedByte ([MS-ASDTYPE] section 2.8) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R71");

                // Verify MS-ASPROV requirement: MS-ASPROV_R71
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    71,
                    @"[In AllowBluetooth] The AllowBluetooth element cannot have child elements.");

                this.VerifyUnsignedByteStructure(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.AllowBluetooth);
            }

            if (policiesSetting.ContainsKey("AllowBrowser"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R80");

                // Verify MS-ASPROV requirement: MS-ASPROV_R80
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    80,
                    @"[In AllowBrowser] The AllowBrowser element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R81");

                // Verify MS-ASPROV requirement: MS-ASPROV_R81
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    81,
                    @"[In AllowBrowser] The AllowBrowser element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowCamera"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R85");

                // Verify MS-ASPROV requirement: MS-ASPROV_R85
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    85,
                    @"[In AllowCamera] The AllowCamera element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R86");

                // Verify MS-ASPROV requirement: MS-ASPROV_R86
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    86,
                    @"[In AllowCamera] The AllowCamera element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowConsumerEmail"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R92");

                // Verify MS-ASPROV requirement: MS-ASPROV_R92
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    92,
                    @"[In AllowConsumerEmail] The AllowConsumerEmail element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R93");

                // Verify MS-ASPROV requirement: MS-ASPROV_R93
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    93,
                    @"[In AllowConsumerEmail] The AllowConsumerEmail element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowDesktopSync"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R97");

                // Verify MS-ASPROV requirement: MS-ASPROV_R97
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    97,
                    @"[In AllowDesktopSync] The AllowDesktopSync element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R98");

                // Verify MS-ASPROV requirement: MS-ASPROV_R98
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    98,
                    @"[In AllowDesktopSync] The AllowDesktopSync element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowHTMLEmail"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R104");

                // Verify MS-ASPROV requirement: MS-ASPROV_R104
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    104,
                    @"[In AllowHTMLEmail] The AllowHTMLEmail element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R105");

                // Verify MS-ASPROV requirement: MS-ASPROV_R105
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    105,
                    @"[In AllowHTMLEmail] The AllowHTMLEmail element cannot have child elements.");

                this.VerifyBooleanStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R106");

                // Verify MS-ASPROV requirement: MS-ASPROV_R106
                // Since MS-ASDTYPE_R5 has been verified in VerifyBooleanStructure, so this requirement can be captured.
                Site.CaptureRequirement(
                    106,
                    @"[In AllowHTMLEmail] Valid values [0,1] for AllowHTMLEmail are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("AllowInternetSharing"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R109");

                // Verify MS-ASPROV requirement: MS-ASPROV_R109
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    109,
                    @"[In AllowInternetSharing] The AllowInternetSharing element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R110");

                // Verify MS-ASPROV requirement: MS-ASPROV_R110
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    110,
                    @"[In AllowInternetSharing] The AllowInternetSharing element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowIrDA"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R116");

                // Verify MS-ASPROV requirement: MS-ASPROV_R116
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    116,
                    @"[In AllowIrDA] The AllowIrDA element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R117");

                // Verify MS-ASPROV requirement: MS-ASPROV_R117
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    117,
                    @"[In AllowIrDA] The AllowIrDA element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowPOPIMAPEmail"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R123");

                // Verify MS-ASPROV requirement: MS-ASPROV_R123
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    123,
                    @"[In AllowPOPIMAPEmail] The AllowPOPIMAPEmail element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R124");

                // Verify MS-ASPROV requirement: MS-ASPROV_R124
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    124,
                    @"[In AllowPOPIMAPEmail] The AllowPOPIMAPEmail element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowRemoteDesktop"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R128");

                // Verify MS-ASPROV requirement: MS-ASPROV_R128
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    128,
                    @"[In AllowRemoteDesktop] The AllowRemoteDesktop element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R129");

                // Verify MS-ASPROV requirement: MS-ASPROV_R129
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    129,
                    @"[In AllowRemoteDesktop] The AllowRemoteDesktop element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowSimpleDevicePassword"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R135");

                // Verify MS-ASPROV requirement: MS-ASPROV_R135
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    135,
                    @"[In AllowSimpleDevicePassword] The AllowSimpleDevicePassword element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R137");

                // Verify MS-ASPROV requirement: MS-ASPROV_R137
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    137,
                    @"[In AllowSimpleDevicePassword] The AllowSimpleDevicePassword element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowSMIMEEncryptionAlgorithmNegotiation"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R143");

                // Verify MS-ASPROV requirement: MS-ASPROV_R143
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    143,
                    @"[In AllowSMIMEEncryptionAlgorithmNegotiation] The AllowSMIMEEncryptionAlgorithmNegotation element is an optional child element of type integer ([MS-ASDTYPE] section 2.6) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R144");

                // Verify MS-ASPROV requirement: MS-ASPROV_R144
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    144,
                    @"[In AllowSMIMEEncryptionAlgorithmNegotiation] The AllowSMIMEEncryptionAlgorithmNegotation element cannot have child elements.");

                this.VerifyIntegerStructure();
            }

            if (policiesSetting.ContainsKey("AllowSMIMESoftCerts"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R149");

                // Verify MS-ASPROV requirement: MS-ASPROV_R149
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    149,
                    @"[In AllowSMIMESoftCerts] The AllowSMIMESoftCerts element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R150");

                // Verify MS-ASPROV requirement: MS-ASPROV_R150
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    150,
                    @"[In AllowSMIMESoftCerts] The AllowSMIMESoftCerts element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowStorageCard"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R154");

                // Verify MS-ASPROV requirement: MS-ASPROV_R154
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    154,
                    @"[In AllowStorageCard] The AllowStorageCard element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R155");

                // Verify MS-ASPROV requirement: MS-ASPROV_R155
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    155,
                    @"[In AllowStorageCard] The AllowStorageCard element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowTextMessaging"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R160");

                // Verify MS-ASPROV requirement: MS-ASPROV_R160
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    160,
                    @"[In AllowTextMessaging] The AllowTextMessaging element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R161");

                // Verify MS-ASPROV requirement: MS-ASPROV_R161
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    161,
                    @"[In AllowTextMessaging] The AllowTextMessaging element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowUnsignedApplications"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R166");

                // Verify MS-ASPROV requirement: MS-ASPROV_R166
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    166,
                    @"[In AllowUnsignedApplications] The AllowUnsignedApplications element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R167");

                // Verify MS-ASPROV requirement: MS-ASPROV_R167
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    167,
                    @"[In AllowUnsignedApplications] The AllowUnsignedApplications element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowUnsignedInstallationPackages"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R171");

                // Verify MS-ASPROV requirement: MS-ASPROV_R171
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    171,
                    @"[In AllowUnsignedInstallationPackages] The AllowUnsignedInstallationPackages element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R172");

                // Verify MS-ASPROV requirement: MS-ASPROV_R172
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    172,
                    @"[In AllowUnsignedInstallationPackages] The AllowUnsignedInstallationPackages element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AllowWiFi"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R176");

                // Verify MS-ASPROV requirement: MS-ASPROV_R176
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    176,
                    @"[In AllowWifi] The AllowWifi element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R177");

                // Verify MS-ASPROV requirement: MS-ASPROV_R177
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    177,
                    @"[In AllowWifi] The AllowWifi element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("AlphanumericDevicePasswordRequired"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R182");

                // Verify MS-ASPROV requirement: MS-ASPROV_R182
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    182,
                    @"[In AlphanumericDevicePasswordRequired] The AlphanumericDevicePasswordRequired element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R183");

                // Verify MS-ASPROV requirement: MS-ASPROV_R183
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    183,
                    @"[In AlphanumericDevicePasswordRequired] The AlphanumericDevicePasswordRequired element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("ApprovedApplicationList"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R194");

                // Verify MS-ASPROV requirement: MS-ASPROV_R194
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    194,
                    @"[In ApprovedApplicationList] The ApprovedApplicationList element is an optional container ([MS-ASDTYPE] section 2.2) element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R195");

                // Verify MS-ASPROV requirement: MS-ASPROV_R195
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    195,
                    @"[In ApprovedApplicationList] It [ApprovedApplicationList element element] is a child of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R199");

                // Verify MS-ASPROV requirement: MS-ASPROV_R199
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    199,
                    @"[In ApprovedApplicationList] A command response has a maximum of one ApprovedApplicationList element per EASProvisionDoc element.");

                this.VerifyContainerStructure();

                if (provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.ApprovedApplicationList.Length != 0)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R202");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R202
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        202,
                        @"[In ApprovedApplicationList] The ApprovedApplicationList element has only the following child element: Hash (section 2.2.2.29): This element is optional.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R236");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R236
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        236,
                        @"[In Hash] The Hash element is an optional child element of type string ([MS-ASDTYPE] section 2.7) of the ApprovedApplicationList element (section 2.2.2.22).");
                }
            }

            if (policiesSetting.ContainsKey("AttachmentsEnabled"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R203");

                // Verify MS-ASPROV requirement: MS-ASPROV_R203
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    203,
                    @"[In AttachmentsEnabled] The AttachmentsEnabled element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R204");

                // Verify MS-ASPROV requirement: MS-ASPROV_R204
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    204,
                    @"[In AttachmentsEnabled] The AttachmentsEnabled element cannot have child elements.");

                this.VerifyBooleanStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R205");

                // Verify MS-ASPROV requirement: MS-ASPROV_R205
                // Since MS-ASDTYPE_R5 has been verified in VerifyBooleanStructure, so this requirement can be captured.
                Site.CaptureRequirement(
                    205,
                    @"[In AttachmentsEnabled] Valid values [0,1] for AttachmentsEnabled are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("DevicePasswordEnabled"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R211");

                // Verify MS-ASPROV requirement: MS-ASPROV_R211
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    211,
                    @"[In DevicePasswordEnabled] The DevicePasswordEnabled element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R212");

                // Verify MS-ASPROV requirement: MS-ASPROV_R212
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    212,
                    @"[In DevicePasswordEnabled] The DevicePasswordEnabled element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("DevicePasswordExpiration"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R216");

                // Verify MS-ASPROV requirement: MS-ASPROV_R216
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    216,
                    @"[In DevicePasswordExpiration] The DevicePasswordExpiration element is an optional child element of type unsignedIntOrEmpty (section 2.2.3.3) of the EASProvisionDoc element, as specified in section 2.2.2.28.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R218");

                // Verify MS-ASPROV requirement: MS-ASPROV_R218
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    218,
                    @"[In DevicePasswordExpiration] The DevicePasswordExpiration element cannot have child elements.");

                if (string.IsNullOrEmpty(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.DevicePasswordExpiration))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R674");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R674
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        674,
                        @"[In unsignedIntOrEmpty Simple Type] The unsignedIntOrEmpty simple type represents a value that can either be [an xs:unsignedInt type, as specified in [XMLSCHEMA2/2] section 3.3.22, or] an empty value.
<xs:simpleType name=""unsignedIntOrEmpty"">
  <xs:union memberTypes=""xs:unsignedInt EmptyVal""/>
</xs:simpleType>");
                }
                else
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R677");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R677
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        677,
                        @"[In unsignedIntOrEmpty Simple Type] The unsignedIntOrEmpty simple type represents a value that can either be an xs:unsignedInt type, as specified in [XMLSCHEMA2/2] section 3.3.22, [or an empty value].
<xs:simpleType name=""unsignedIntOrEmpty"">
  <xs:union memberTypes=""xs:unsignedInt EmptyVal""/>
</xs:simpleType>");
                }
            }

            if (policiesSetting.ContainsKey("DevicePasswordHistory"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R224");

                // Verify MS-ASPROV requirement: MS-ASPROV_R224
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    224,
                    @"[In DevicePasswordHistory] The DevicePasswordHistory element is an optional child element of type unsignedInt of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R225");

                // Verify MS-ASPROV requirement: MS-ASPROV_R225
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    225,
                    @"[In DevicePasswordHistory] The DevicePasswordHistory element cannot have child elements.");
            }

            if (policiesSetting.ContainsKey("MaxAttachmentSize"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R240");

                // Verify MS-ASPROV requirement: MS-ASPROV_R240
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    240,
                    @"[In MaxAttachmentSize] The MaxAttachmentSize element is an optional child element of type unsignedIntOrEmpty (section 2.2.3.3) of the EASProvisionDoc element, as specified in section 2.2.2.28.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R241");

                // Verify MS-ASPROV requirement: MS-ASPROV_R241
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    241,
                    @"[In MaxAttachmentSize] The EASProvisionDoc element has at most one instance of the MaxAttachmentSize element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R242");

                // Verify MS-ASPROV requirement: MS-ASPROV_R242
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    242,
                    @"[In MaxAttachmentSize] The MaxAttachmentSize element cannot have child elements.");
            }

            if (policiesSetting.ContainsKey("MaxCalendarAgeFilter"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R243");

                // Verify MS-ASPROV requirement: MS-ASPROV_R243
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    243,
                    @"[In MaxCalendarAgeFilter] The MaxCalendarAgeFilter element is an optional child element of type unsignedInt ([XMLSCHEMA2/2] section 3.3.22) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R244");

                // Verify MS-ASPROV requirement: MS-ASPROV_R244
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    244,
                    @"[In MaxCalendarAgeFilter] The MaxCalendarAgeFilter element cannot have child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R245");
                Common.VerifyActualValues("MaxCalendarAgeFilter", new string[] { "0", "4", "5", "6", "7" }, provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxCalendarAgeFilter.ToString(), Site);

                // Verify MS-ASPROV requirement: MS-ASPROV_R245
                // The value of MaxCalendarAgeFilter element is one of the valid values, so this requirement can be captured.
                Site.CaptureRequirement(
                    245,
                    @"[In MaxCalendarAgeFilter] Valid values [0,4,5,6,7] for MaxCalendarAgeFilter are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("MaxDevicePasswordFailedAttempts"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R251");

                // Verify MS-ASPROV requirement: MS-ASPROV_R251
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    251,
                    @"[In MaxDevicePasswordFailedAttempts] The MaxDevicePasswordFailedAttempts element is an optional child element of type unsignedByteOrEmpty (section 2.2.3.2) of the EASProvisionDoc element, as specified in section 2.2.2.28.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R252");

                // Verify MS-ASPROV requirement: MS-ASPROV_R252
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    252,
                    @"[In MaxDevicePasswordFailedAttempts] The MaxDevicePasswordFailedAttempts element cannot have child elements.");

                if (string.IsNullOrEmpty(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxDevicePasswordFailedAttempts))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R673");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R673
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        673,
                        @"[In unsignedByteOrEmpty Simple Type] The unsignedByteOrEmpty simple type represents a value that can either be an [xs:unsignedByte type, as specified in [XMLSCHEMA2/2] section 3.3.24, or] an empty value.
<xs:simpleType name=""unsignedByteOrEmpty"">
  <xs:union memberTypes=""xs:unsignedByte EmptyVal""/>
</xs:simpleType>");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R672");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R672
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        672,
                        @"[In EmptyVal Simple Type] The EmptyVal simple type represents an empty value.
<xs:simpleType name=""EmptyVal"">
  <xs:restriction base=""xs:string"">
    <xs:maxLength value=""0""/>
  </xs:restriction>
</xs:simpleType>");
                }
                else
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R676");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R676
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        676,
                        @"[In unsignedByteOrEmpty Simple Type] The unsignedByteOrEmpty simple type represents a value that can either be an xs:unsignedByte type, as specified in [XMLSCHEMA2/2] section 3.3.24, [or an empty value].
<xs:simpleType name=""unsignedByteOrEmpty"">
  <xs:union memberTypes=""xs:unsignedByte EmptyVal""/>
</xs:simpleType>");
                }
            }

            if (policiesSetting.ContainsKey("MaxEmailAgeFilter"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R255");

                // Verify MS-ASPROV requirement: MS-ASPROV_R255
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    255,
                    @"[In MaxEmailAgeFilter] The MaxEmailAgeFilter element is an optional child element of type unsignedInt ([XMLSCHEMA2/2] section 3.3.22) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R256");

                // Verify MS-ASPROV requirement: MS-ASPROV_R256
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    256,
                    @"[In MaxEmailAgeFilter] The MaxEmailAgeFilter element cannot have child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R257");

                Common.VerifyActualValues("MaxEmailAgeFilter", new string[] { "0", "1", "2", "3", "4", "5" }, provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxEmailAgeFilter.ToString(), Site);

                // Verify MS-ASPROV requirement: MS-ASPROV_R257
                // The value of MaxEmailAgeFilter element is one of the valid values, so this requirement can be captured.
                Site.CaptureRequirement(
                    257,
                    @"[In MaxEmailAgeFilter] Valid values [0,1,2,3,4,5] are listed in the following table and represent the maximum allowable number of days to sync email.");
            }

            if (policiesSetting.ContainsKey("MaxEmailBodyTruncationSize"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R265");

                // Verify MS-ASPROV requirement: MS-ASPROV_R265
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    265,
                    @"[In MaxEmailBodyTruncationSize] The MaxEmailBodyTruncationSize element cannot have child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R266, the value of MaxEmailBodyTruncationSize element is {0}.", provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxEmailBodyTruncationSize);

                // Verify MS-ASPROV requirement: MS-ASPROV_R266
                bool isVerifiedR266 = Convert.ToInt32(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxEmailBodyTruncationSize) >= -1;

                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR266,
                    266,
                    @"[In MaxEmailBodyTruncationSize] Valid values [-1, 0, >0] for the MaxEmailBodyTruncationSize element are an integer ([MS-ASDTYPE] section 2.6) of one of the values or ranges listed in the following table.");

                this.VerifyIntegerStructure();
            }

            if (policiesSetting.ContainsKey("MaxEmailHTMLBodyTruncationSize"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R271");

                // Verify MS-ASPROV requirement: MS-ASPROV_R271
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    271,
                    @"[In MaxEmailHTMLBodyTruncationSize] The MaxEmailHTMLBodyTruncationSize element cannot have child elements.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R272, the value of MaxEmailHTMLBodyTruncationSize element is {0}.", provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxEmailHTMLBodyTruncationSize);

                // Verify MS-ASPROV requirement: MS-ASPROV_R272
                bool isVerifiedR272 = Convert.ToInt32(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MaxEmailHTMLBodyTruncationSize) >= -1;

                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR272,
                    272,
                    @"[In MaxEmailHTMLBodyTruncationSize] Valid values [-1, 0, >0] for the MaxEmailHTMLBodyTruncationSize element are an integer ([MS-ASDTYPE] section 2.6) of one of the values or ranges listed in the following table.");

                this.VerifyIntegerStructure();
            }

            if (policiesSetting.ContainsKey("MaxInactivityTimeDeviceLock"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R276");

                // Verify MS-ASPROV requirement: MS-ASPROV_R276
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    276,
                    @"[In MaxInactivityTimeDeviceLock] The MaxInactivityTimeDeviceLock element is an optional child element of type unsignedIntOrEmpty (section 2.2.3.3) of the EASProvisionDoc element, as specified in section 2.2.2.28.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R277");

                // Verify MS-ASPROV requirement: MS-ASPROV_R277
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    277,
                    @"[In MaxInactivityTimeDeviceLock] The MaxInactivityTimeDeviceLock element cannot have child elements.");
            }

            if (policiesSetting.ContainsKey("MinDevicePasswordComplexCharacters"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R280");

                // Verify MS-ASPROV requirement: MS-ASPROV_R280
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    280,
                    @"[In MinDevicePasswordComplexCharacters] The MinDevicePasswordComplexCharacters element is an optional child element of type unsignedByte ([MS-ASDTYPE] section 2.8) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R281");

                // Verify MS-ASPROV requirement: MS-ASPROV_R281
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    281,
                    @"[In MinDevicePasswordComplexCharacters] The MinDevicePasswordComplexCharacters element cannot have child elements.");

                this.VerifyUnsignedByteStructure(provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.MinDevicePasswordComplexCharacters);
            }

            if (policiesSetting.ContainsKey("MinDevicePasswordLength"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R286");

                // Verify MS-ASPROV requirement: MS-ASPROV_R286
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    286,
                    @"[In MinDevicePasswordLength] The MinDevicePasswordLength element is an optional child element of type unsignedByteOrEmpty (section 2.2.3.2) of the EASProvisionDoc element, as specified in section 2.2.2.28.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R287");

                // Verify MS-ASPROV requirement: MS-ASPROV_R287
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    287,
                    @"[In MinDevicePasswordLength] The MinDevicePasswordLength element cannot have child elements.");
            }

            if (policiesSetting.ContainsKey("PasswordRecoveryEnabled"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R291");

                // Verify MS-ASPROV requirement: MS-ASPROV_R291
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    291,
                    @"[In PasswordRecoveryEnabled] The PasswordRecoveryEnabled element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R292");

                // Verify MS-ASPROV requirement: MS-ASPROV_R292
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    292,
                    @"[In PasswordRecoveryEnabled] The PasswordRecoveryEnabled element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("RequireDeviceEncryption"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R326");

                // Verify MS-ASPROV requirement: MS-ASPROV_R326
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    326,
                    @"[In RequireDeviceEncryption] The RequireDeviceEncryption element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R327");

                // Verify MS-ASPROV requirement: MS-ASPROV_R327
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    327,
                    @"[In RequireDeviceEncryption] The RequireDeviceEncryption element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("RequireEncryptedSMIMEMessages"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R331");

                // Verify MS-ASPROV requirement: MS-ASPROV_R331
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    331,
                    @"[In RequireEncryptedSMIMEMessages] The RequireEncryptedSMIMEMessages element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R327");

                // Verify MS-ASPROV requirement: MS-ASPROV_R327
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    332,
                    @"[In RequireEncryptedSMIMEMessages] The RequireEncryptedSMIMEMessages element cannot have child elements.");

                this.VerifyBooleanStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R333");

                // Verify MS-ASPROV requirement: MS-ASPROV_R333
                // Since MS-ASDTYPE_R5 has been verified in VerifyBooleanStructure, so this requirement can be captured.
                Site.CaptureRequirement(
                    333,
                    @"[In RequireEncryptedSMIMEMessages] Valid values [0,1] for RequireEncryptedSMIMEMessages are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("RequireEncryptionSMIMEAlgorithm"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R336");

                // Verify MS-ASPROV requirement: MS-ASPROV_R336
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    336,
                    @"[In RequireEncryptionSMIMEAlgorithm] The RequireEncryptionSMIMEAlgorithm element is an optional child element of type integer ([MS-ASDTYPE] section 2.6) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R337");

                // Verify MS-ASPROV requirement: MS-ASPROV_R337
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    337,
                    @"[In RequireEncryptionSMIMEAlgorithm] The RequireEncryptionSMIMEAlgorithm element cannot have child elements.");

                this.VerifyIntegerStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R338");

                Common.VerifyActualValues("RequireEncryptionSMIMEAlgorithm", new string[] { "0", "1", "2", "3", "4" }, provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.RequireEncryptionSMIMEAlgorithm.ToString(), Site);

                // Verify MS-ASPROV requirement: MS-ASPROV_R338
                // The actual value is one of the valid values, so this requirement can be captured.
                Site.CaptureRequirement(
                    338,
                    @"[In RequireEncryptionSMIMEAlgorithm] Valid values [0,1,2,3,4] for RequireEncryptionSMIMEAlgorithm are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("RequireManualSyncWhenRoaming"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R344");

                // Verify MS-ASPROV requirement: MS-ASPROV_R344
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    344,
                    @"[In RequireManualSyncWhenRoaming] The RequireManualSyncWhenRoaming element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R345");

                // Verify MS-ASPROV requirement: MS-ASPROV_R345
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    345,
                    @"[In RequireManualSyncWhenRoaming] The RequireManualSyncWhenRoaming element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("RequireSignedSMIMEAlgorithm"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R349");

                // Verify MS-ASPROV requirement: MS-ASPROV_R349
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    349,
                    @"[In RequireSignedSMIMEAlgorithm] The RequireSignedSMIMEAlgorithm element is an optional child element of type integer ([MS-ASDTYPE] section 2.6) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R350");

                // Verify MS-ASPROV requirement: MS-ASPROV_R350
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    350,
                    @"[In RequireSignedSMIMEAlgorithm] The RequireSignedSMIMEAlgorithm element cannot have child elements.");

                this.VerifyIntegerStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R351");

                Common.VerifyActualValues("RequireSignedSMIMEAlgorithm", new string[] { "0", "1" }, provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.RequireSignedSMIMEAlgorithm.ToString(), Site);

                // Verify MS-ASPROV requirement: MS-ASPROV_R351
                // The value of RequireSignedSMIMEAlgorithm element is one of valid values, so this requirement can be captured.
                Site.CaptureRequirement(
                    351,
                    @"[In RequireSignedSMIMEAlgorithm] Valid values [0,1] for RequireSignedSMIMEAlgorithm are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("RequireSignedSMIMEMessages"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R354");

                // Verify MS-ASPROV requirement: MS-ASPROV_R354
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    354,
                    @"[In RequireSignedSMIMEMessages] The RequireSignedSMIMEMessages element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R355");

                // Verify MS-ASPROV requirement: MS-ASPROV_R355
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    355,
                    @"[In RequireSignedSMIMEMessages] The RequireSignedSMIMEMessages element cannot have child elements.");

                this.VerifyBooleanStructure();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R356");

                // Verify MS-ASPROV requirement: MS-ASPROV_R356
                // Since MS-ASDTYPE_R5 has been verified in VerifyBooleanStructure, so this requirement can be captured.
                Site.CaptureRequirement(
                    356,
                    @"[In RequireSignedSMIMEMessages] Valid values [0,1] for RequireSignedSMIMEMessages are listed in the following table.");
            }

            if (policiesSetting.ContainsKey("RequireStorageCardEncryption"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R359");

                // Verify MS-ASPROV requirement: MS-ASPROV_R359
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    359,
                    @"[In RequireStorageCardEncryption] The RequireStorageCardEncryption element is an optional child element of type boolean ([MS-ASDTYPE] section 2.1) of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R360");

                // Verify MS-ASPROV requirement: MS-ASPROV_R360
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    360,
                    @"[In RequireStorageCardEncryption] The RequireStorageCardEncryption element cannot have child elements.");

                this.VerifyBooleanStructure();
            }

            if (policiesSetting.ContainsKey("UnapprovedInROMApplicationList"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R404");

                // Verify MS-ASPROV requirement: MS-ASPROV_R404
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    404,
                    @"[In UnapprovedInROMApplicationList] The UnapprovedInROMApplicationList element is an optional container ([MS-ASDTYPE] section 2.2) element.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R405");

                // Verify MS-ASPROV requirement: MS-ASPROV_R405
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    405,
                    @"[In UnapprovedInROMApplicationList] It [UnapprovedInROMApplicationList element] is a child of the EASProvisionDoc element (section 2.2.2.28).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R408");

                // Verify MS-ASPROV requirement: MS-ASPROV_R408
                // The schema has been validated, so this requirement can be captured.
                Site.CaptureRequirement(
                    408,
                    @"[In UnapprovedInROMApplicationList] A command response has a maximum of one UnapprovedInROMApplicationList element per EASProvisionDoc element.");

                this.VerifyContainerStructure();

                if (provisionResponse.ResponseData.Policies.Policy.Data.EASProvisionDoc.UnapprovedInROMApplicationList.Length != 0)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R189");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R189
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        189,
                        @"[In ApplicationName] The ApplicationName element is an optional child element of type string ([MS-ASDTYPE] section 2.7) of the UnapprovedInROMApplicationList element (section 2.2.2.55).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASPROV_R409");

                    // Verify MS-ASPROV requirement: MS-ASPROV_R409
                    // The schema has been validated, so this requirement can be captured.
                    Site.CaptureRequirement(
                        409,
                        @"[In UnapprovedInROMApplicationList] The UnapprovedInROMApplicationList element has only the following child element: ApplicationName (section 2.2.2.21): This element is optional.");
                }
            }
        }
        #endregion

        #region Verify data structure
        /// <summary>
        /// This method is used to verify the Boolean related requirements.
        /// </summary>
        private void VerifyBooleanStructure()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R4");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R4
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                4,
                @"[In boolean Data Type] It [a boolean] is declared as an element with a type attribute of ""boolean"".");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R5");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R5
            Site.CaptureRequirement(
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
        /// This method is used to verify the Container related requirements.
        /// </summary>
        private void VerifyContainerStructure()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R8");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R8
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                8,
                @"[In container Data Type] A container is an XML element that encloses other elements but has no value of its own.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R9");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R9
            // The schema has been validated, so this requirement can be captured.
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                9,
                @"[In container Data Type] It [container] is a complex type with complex content, as specified in [XMLSCHEMA1/2] section 3.4.2.");
        }

        /// <summary>
        /// This method is used to verify the integer related requirements.
        /// </summary>
        private void VerifyIntegerStructure()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R87");

            // ActiveSyncClient encoded integer data as inline strings, so if response is successfully returned this requirement can be covered.
            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R87
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                87,
                @"[In integer Data Type] Elements with an integer data type MUST be encoded and transmitted as WBXML inline strings, as specified in [WBXML1.2].");
        }

        /// <summary>
        /// This method is used to verify the string related requirements.
        /// </summary>
        private void VerifyStringStructure()
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R88");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R88
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                88,
                @"[In string Data Type] A string is a chunk of Unicode text.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASDTYPE_R90");

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R90
            Site.CaptureRequirement(
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
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                94,
                @"[In string Data Type] Elements of these types [ActiveSync defines several conventions for strings that adhere to commonly used formats]are defined as string types in XML schemas.");
        }

        /// <summary>
        /// This method is used to verify the unsignedByte related requirements.
        /// </summary>
        /// <param name="byteValue">A byte value.</param>
        private void VerifyUnsignedByteStructure(byte? byteValue)
        {
            // If the schema validation is successful, then following requirements can be captured.
            Site.Assert.IsTrue(this.activeSyncClient.ValidationResult, "Schema validation should be successful.");

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

            // Verify MS-ASDTYPE requirement: MS-ASDTYPE_R125
            Site.CaptureRequirement(
                "MS-ASDTYPE",
                125,
                @"[In unsignedByte Data Type] Elements of this type [unsignedByte type] are declared with an element whose type attribute is set to ""unsignedByte"".");
        }

        #endregion

        #region Verify code page 14 requirements of [MS-ASWBXML].
        /// <summary>
        /// Verify WBXML Capture for WBXML process.
        /// </summary>
        private void VerifyWBXMLRequirements()
        {
            // Get WBXML decoded data.
            Dictionary<string, int> decodedData = this.activeSyncClient.GetMSASWBXMLImplementationInstance().DecodeDataCollection;

            // Find Code Page 14.
            foreach (KeyValuePair<string, int> decodeDataItem in decodedData)
            {
                byte token;
                string tagName = Common.GetTagName(decodeDataItem.Key, out token);
                string codePageName = Common.GetCodePageName(decodeDataItem.Key);
                int codepage = decodeDataItem.Value;
                bool isValidCodePage = codepage >= 0 && codepage <= 24;
                Site.Assert.IsTrue(isValidCodePage, "Code page value should between 0-24, the actual value is :{0}", codepage);

                // Capture requirements.
                if (14 == codepage)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R24");

                    // Verify MS-ASWBXML requirement: MS-ASWBXML_R24
                    Site.CaptureRequirementIfAreEqual<string>(
                        "provision",
                        codePageName.ToLower(System.Globalization.CultureInfo.CurrentCulture),
                        "MS-ASWBXML",
                        24,
                        @"[In Code Pages] [This algorithm supports] [Code page] 14 [that indicates] [XML namespace] Provision");

                    switch (tagName)
                    {
                        case "Provision":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R337");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R337
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x05,
                                    token,
                                    "MS-ASWBXML",
                                    337,
                                    @"[In Code Page 14: Provision] [Tag name] Provision [Token] 0x05 [supports protocol versions] All");

                                break;
                            }

                        case "Policies":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R338");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R338
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x06,
                                    token,
                                    "MS-ASWBXML",
                                    338,
                                    @"[In Code Page 14: Provision] [Tag name] Policies [Token] 0x06 [supports protocol versions] All");

                                break;
                            }

                        case "Policy":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R339");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R339
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x07,
                                    token,
                                    "MS-ASWBXML",
                                    339,
                                    @"[In Code Page 14: Provision] [Tag name] Policy [Token] 0x07 [supports protocol versions] All");

                                break;
                            }

                        case "PolicyType":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R340");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R340
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x08,
                                    token,
                                    "MS-ASWBXML",
                                    340,
                                    @"[In Code Page 14: Provision] [Tag name] PolicyType [Token] 0x08 [supports protocol versions] All");

                                break;
                            }

                        case "PolicyKey":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R341");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R341
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x09,
                                    token,
                                    "MS-ASWBXML",
                                    341,
                                    @"[In Code Page 14: Provision] [Tag name] PolicyKey [Token] 0x09 [supports protocol versions] All");

                                break;
                            }

                        case "Data":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R342");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R342
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0A,
                                    token,
                                    "MS-ASWBXML",
                                    342,
                                    @"[In Code Page 14: Provision] [Tag name] Data [Token] 0x0A [supports protocol versions] All");

                                break;
                            }

                        case "Status":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R343");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R343
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0B,
                                    token,
                                    "MS-ASWBXML",
                                    343,
                                    @"[In Code Page 14: Provision] [Tag name] Status [Token] 0x0B [supports protocol versions] All");

                                break;
                            }

                        case "RemoteWipe":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R344");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R344
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0C,
                                    token,
                                    "MS-ASWBXML",
                                    344,
                                    @"[In Code Page 14: Provision] [Tag name] RemoteWipe [Token] 0x0C [supports protocol versions] All");

                                break;
                            }

                        case "EASProvisionDoc":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R345");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R345
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0D,
                                    token,
                                    "MS-ASWBXML",
                                    345,
                                    @"[In Code Page 14: Provision] [Tag name] EASProvisionDoc [Token] 0x0D [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "DevicePasswordEnabled":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R346");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R346
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0E,
                                    token,
                                    "MS-ASWBXML",
                                    346,
                                    @"[In Code Page 14: Provision] [Tag name] DevicePasswordEnabled [Token] 0x0E [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AlphanumericDevicePasswordRequired":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R347");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R347
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x0F,
                                    token,
                                    "MS-ASWBXML",
                                    347,
                                    @"[In Code Page 14: Provision] [Tag name] AlphanumericDevicePasswordRequired [Token] 0x0F [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireStorageCardEncryption":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R349");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R349
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x10,
                                    token,
                                    "MS-ASWBXML",
                                    349,
                                    @"[In Code Page 14: Provision] [Tag name] RequireStorageCardEncryption [Token] 0x10 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "PasswordRecoveryEnabled":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R350");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R350
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x11,
                                    token,
                                    "MS-ASWBXML",
                                    350,
                                    @"[In Code Page 14: Provision] [Tag name] PasswordRecoveryEnabled [Token] 0x11 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AttachmentsEnabled":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R351");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R351
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x13,
                                    token,
                                    "MS-ASWBXML",
                                    351,
                                    @"[In Code Page 14: Provision] [Tag name] AttachmentsEnabled [Token] 0x13 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MinDevicePasswordLength":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R352");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R352
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x14,
                                    token,
                                    "MS-ASWBXML",
                                    352,
                                    @"[In Code Page 14: Provision] [Tag name] MinDevicePasswordLength [Token] 0x14 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxInactivityTimeDeviceLock":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R353");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R353
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x15,
                                    token,
                                    "MS-ASWBXML",
                                    353,
                                    @"[In Code Page 14: Provision] [Tag name] MaxInactivityTimeDeviceLock [Token] 0x15 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxDevicePasswordFailedAttempts":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R354");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R354
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x16,
                                    token,
                                    "MS-ASWBXML",
                                    354,
                                    @"[In Code Page 14: Provision] [Tag name] MaxDevicePasswordFailedAttempts [Token] 0x16 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxAttachmentSize":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R355");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R355
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x17,
                                    token,
                                    "MS-ASWBXML",
                                    355,
                                    @"[In Code Page 14: Provision] [Tag name] MaxAttachmentSize [Token] 0x17 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowSimpleDevicePassword":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R356");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R356
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x18,
                                    token,
                                    "MS-ASWBXML",
                                    356,
                                    @"[In Code Page 14: Provision] [Tag name] AllowSimpleDevicePassword [Token] 0x18 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "DevicePasswordExpiration":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R357");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R357
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x19,
                                    token,
                                    "MS-ASWBXML",
                                    357,
                                    @"[In Code Page 14: Provision] [Tag name] DevicePasswordExpiration [Token] 0x19 [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "DevicePasswordHistory":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R358");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R358
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1A,
                                    token,
                                    "MS-ASWBXML",
                                    358,
                                    @"[In Code Page 14: Provision] [Tag name] DevicePasswordHistory [Token] 0x1A [supports protocol versions] 12.0, 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowStorageCard":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R359");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R359
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1B,
                                    token,
                                    "MS-ASWBXML",
                                    359,
                                    @"[In Code Page 14: Provision] [Tag name] AllowStorageCard [Token] 0x1B [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowCamera":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R360");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R360
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1C,
                                    token,
                                    "MS-ASWBXML",
                                    360,
                                    @"[In Code Page 14: Provision] [Tag name] AllowCamera [Token] 0x1C [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireDeviceEncryption":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R361");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R361
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1D,
                                    token,
                                    "MS-ASWBXML",
                                    361,
                                    @"[In Code Page 14: Provision] [Tag name] RequireDeviceEncryption [Token] 0x1D [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowUnsignedApplications":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R362");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R362
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1E,
                                    token,
                                    "MS-ASWBXML",
                                    362,
                                    @"[In Code Page 14: Provision] [Tag name] AllowUnsignedApplications [Token] 0x1E [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowUnsignedInstallationPackages":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R363");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R363
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x1F,
                                    token,
                                    "MS-ASWBXML",
                                    363,
                                    @"[In Code Page 14: Provision] [Tag name] AllowUnsignedInstallationPackages [Token] 0x1F [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MinDevicePasswordComplexCharacters":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R364");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R364
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x20,
                                    token,
                                    "MS-ASWBXML",
                                    364,
                                    @"[In Code Page 14: Provision] [Tag name] MinDevicePasswordComplexCharacters [Token] 0x20 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowWiFi":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R365");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R365
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x21,
                                    token,
                                    "MS-ASWBXML",
                                    365,
                                    @"[In Code Page 14: Provision] [Tag name] AllowWiFi [Token] 0x21 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowTextMessaging":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R366");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R366
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x22,
                                    token,
                                    "MS-ASWBXML",
                                    366,
                                    @"[In Code Page 14: Provision] [Tag name] AllowTextMessaging [Token] 0x22 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowPOPIMAPEmail":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R367");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R367
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x23,
                                    token,
                                    "MS-ASWBXML",
                                    367,
                                    @"[In Code Page 14: Provision] [Tag name] AllowPOPIMAPEmail [Token] 0x23 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowBluetooth":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R368");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R368
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x24,
                                    token,
                                    "MS-ASWBXML",
                                    368,
                                    @"[In Code Page 14: Provision] [Tag name] AllowBluetooth [Token] 0x24 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowIrDA":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R369");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R369
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x25,
                                    token,
                                    "MS-ASWBXML",
                                    369,
                                    @"[In Code Page 14: Provision] [Tag name]AllowIrDA [Token] 0x25 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireManualSyncWhenRoaming":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R370");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R370
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x26,
                                    token,
                                    "MS-ASWBXML",
                                    370,
                                    @"[In Code Page 14: Provision] [Tag name] RequireManualSyncWhenRoaming [Token] 0x26 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowDesktopSync":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R371");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R371
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x27,
                                    token,
                                    "MS-ASWBXML",
                                    371,
                                    @"[In Code Page 14: Provision] [Tag name] AllowDesktopSync [Token] 0x27 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxCalendarAgeFilter":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R372");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R372
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x28,
                                    token,
                                    "MS-ASWBXML",
                                    372,
                                    @"[In Code Page 14: Provision] [Tag name] MaxCalendarAgeFilter [Token] 0x28 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowHTMLEmail":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R373");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R373
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x29,
                                    token,
                                    "MS-ASWBXML",
                                    373,
                                    @"[In Code Page 14: Provision] [Tag name] AllowHTMLEmail [Token] 0x29 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxEmailAgeFilter":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R374");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R374
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2A,
                                    token,
                                    "MS-ASWBXML",
                                    374,
                                    @"[In Code Page 14: Provision] [Tag name] MaxEmailAgeFilter [Token] 0x2A [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxEmailBodyTruncationSize":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R375");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R375
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2B,
                                    token,
                                    "MS-ASWBXML",
                                    375,
                                    @"[In Code Page 14: Provision] [Tag name] MaxEmailBodyTruncationSize [Token] 0x2B [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "MaxEmailHTMLBodyTruncationSize":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R376");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R376
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2C,
                                    token,
                                    "MS-ASWBXML",
                                    376,
                                    @"[In Code Page 14: Provision] [Tag name] MaxEmailHTMLBodyTruncationSize [Token] 0x2C [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireSignedSMIMEMessages":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R377");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R377
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2D,
                                    token,
                                    "MS-ASWBXML",
                                    377,
                                    @"[In Code Page 14: Provision] [Tag name] RequireSignedSMIMEMessages [Token] 0x2D [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireEncryptedSMIMEMessages":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R378");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R378
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2E,
                                    token,
                                    "MS-ASWBXML",
                                    378,
                                    @"[In Code Page 14: Provision] [Tag name] RequireEncryptedSMIMEMessages [Token] 0x2E [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireSignedSMIMEAlgorithm":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R379");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R379
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x2F,
                                    token,
                                    "MS-ASWBXML",
                                    379,
                                    @"[In Code Page 14: Provision] [Tag name] RequireSignedSMIMEAlgorithm [Token] 0x2F [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "RequireEncryptionSMIMEAlgorithm":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R380");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R380
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x30,
                                    token,
                                    "MS-ASWBXML",
                                    380,
                                    @"[In Code Page 14: Provision] [Tag name] RequireEncryptionSMIMEAlgorithm [Token] 0x30 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowSMIMEEncryptionAlgorithmNegotiation":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R381");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R381
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x31,
                                    token,
                                    "MS-ASWBXML",
                                    381,
                                    @"[In Code Page 14: Provision] [Tag name] AllowSMIMEEncryptionAlgorithmNegotiation [Token] 0x31 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowSMIMESoftCerts":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R382");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R382
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x32,
                                    token,
                                    "MS-ASWBXML",
                                    382,
                                    @"[In Code Page 14: Provision] [Tag name] AllowSMIMESoftCerts [Token] 0x32 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowBrowser":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R383");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R383
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x33,
                                    token,
                                    "MS-ASWBXML",
                                    383,
                                    @"[In Code Page 14: Provision] [Tag name] AllowBrowser [Token] 0x33 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowConsumerEmail":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R384");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R384
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x34,
                                    token,
                                    "MS-ASWBXML",
                                    384,
                                    @"[In Code Page 14: Provision] [Tag name] AllowConsumerEmail [Token] 0x34 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowRemoteDesktop":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R385");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R385
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x35,
                                    token,
                                    "MS-ASWBXML",
                                    385,
                                    @"[In Code Page 14: Provision] [Tag name] AllowRemoteDesktop [Token] 0x35 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AllowInternetSharing":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R386");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R386
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x36,
                                    token,
                                    "MS-ASWBXML",
                                    386,
                                    @"[In Code Page 14: Provision] [Tag name] AllowInternetSharing [Token] 0x36 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "UnapprovedInROMApplicationList":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R387");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R387
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x37,
                                    token,
                                    "MS-ASWBXML",
                                    387,
                                    @"[In Code Page 14: Provision] [Tag name] UnapprovedInROMApplicationList [Token] 0x37 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "ApplicationName":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R388");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R388
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x38,
                                    token,
                                    "MS-ASWBXML",
                                    388,
                                    @"[In Code Page 14: Provision] [Tag name] ApplicationName [Token] 0x38 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "ApprovedApplicationList":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R389");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R389
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x39,
                                    token,
                                    "MS-ASWBXML",
                                    389,
                                    @"[In Code Page 14: Provision] [Tag name] ApprovedApplicationList [Token] 0x39 [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "Hash":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R390");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R390
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x3A,
                                    token,
                                    "MS-ASWBXML",
                                    390,
                                    @"[In Code Page 14: Provision] [Tag name] Hash [Token] 0x3A [supports protocol versions] 12.1, 14.0, 14.1, 16.0, 16.1");

                                break;
                            }

                        case "AccountOnlyRemoteWipe":
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASWBXML_R3910");

                                // Verify MS-ASWBXML requirement: MS-ASWBXML_R3910
                                Site.CaptureRequirementIfAreEqual<byte>(
                                    0x3B,
                                    token,
                                    "MS-ASWBXML",
                                    3910,
                                    @"[In Code Page 14: Provision] [Tag name] AccountOnlyRemoteWipe [Token] 0x3B [supports protocol versions] 16.1");
                                break;
                            }

                        default:
                            {
                                Site.Assert.Fail("There are unexpected Tag exist in wbxml processing\r\n CodePage[{0}]:TagName[{1}]-Token[0x{2:X}]", codepage, tagName, token);
                                break;
                            }
                    }
                }
            }
        }

        #endregion
    }
}