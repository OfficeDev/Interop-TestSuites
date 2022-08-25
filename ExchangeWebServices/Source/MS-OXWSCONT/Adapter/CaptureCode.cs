namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSCONT.
    /// </summary>
    public partial class MS_OXWSCONTAdapter
    {
        #region Verify requirements related to GetItem operation
        /// <summary>
        /// Capture GetItemResponseType related requirements.
        /// </summary>
        /// <param name="getItemResponse">The response message of GetItem operation.</param>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyGetContactItem(GetItemResponseType getItemResponse, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R114");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R114
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                114,
                @"[In GetItem] The following is the WSDL port type specification for the GetItem operation:<wsdl:operation name=""GetItem"">
                    <wsdl:input message=""tns:GetItemSoapIn"" />
                    <wsdl:output message=""tns:GetItemSoapOut"" />
                    </wsdl:operation>");
     
            ContactItemType[] contacts = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);
            foreach (ContactItemType contact in contacts)
            {
                // Capture ContactItemType Complex Type related requirements.
                this.VerifyContactItemTypeComplexType(contact, isSchemaValidated);                    
            }
        }
        #endregion

        #region Verify requirements related to DeleteItem operation
        /// <summary>
        /// Capture DeleteItemResponseType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyDeleteContactItem(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R274");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R274
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                274,
                @"[In DeleteItem] The following is the WSDL port type specification for the DeleteItem operation: <wsdl:operation name=""DeleteItem"">
                      <wsdl:input message=""tns:DeleteItemSoapIn"" />
                      <wsdl:output message=""tns:DeleteItemSoapOut"" />
                      </wsdl:operation>");
        }
        #endregion

        #region Verify requirements related to UpdateItem operation
        /// <summary>
        /// Capture UpdateItemResponseType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyUpdateContactItem(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R280");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R280
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                280,
                @"[In UpdateItem] The following is the WSDL port type specification for the UpdateItem operation: <wsdl:operation name=""UpdateItem"">
                    <wsdl:input message=""tns:UpdateItemSoapIn"" />
                    <wsdl:output message=""tns:UpdateItemSoapOut"" />
                </wsdl:operation>");
        }
        #endregion

        #region Verify requirements related to MoveItem operation
        /// <summary>
        /// Capture MoveItemResponseType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyMoveContactItem(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R286");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R286
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                286,
                @"[In MoveItem] The following is the WSDL port type specification for the MoveItem operation: <wsdl:operation name=""MoveItem"">
                      <wsdl:input message=""tns:MoveItemSoapIn"" />
                      <wsdl:output message=""tns:MoveItemSoapOut"" />
                      </wsdl:operation>");
        }
        #endregion

        #region Verify requirements related to CopyItem operation
        /// <summary>
        /// Capture CopyItemResponseType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyCopyContactItem(CopyItemResponseType copyItemResponse,bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R292");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R292
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                292,
                @"[In CopyItem] The following is the WSDL port type specification for the CopyItem operation: <wsdl:operation name=""CopyItem"">
                      <wsdl:input message=""tns:CopyItemSoapIn"" />
                      <wsdl:output message=""tns:CopyItemSoapOut"" />
                      </wsdl:operation>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(copyItemResponse);
        }
        #endregion

        #region Verify requirements related to AbchPersonItemType complex types
        /// <summary>
        /// Capture AbchPersonItemType Complex Type related requirements.
        /// </summary>
        /// <param name="abchPersonItemType">A person item from the response package of GetItem operation.</param>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyAbchPersonItemTypeComplexType(AbchPersonItemType abchPersonItemType, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation result should be true!");

            if (Common.IsRequirementEnabled(336002, this.Site) && abchPersonItemType != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16003");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R16003
                // If the abchPersonItemType element is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    16003,
                    @"[In t:AbchPersonItemType Complex Type] The type [AbchPersonItemType] is defined as follow:
  <xs:complexType name=""AbchPersonItemType"" >
   < xs:complexContent >
     < xs:extension base = ""t:ItemType"" >
       < xs:sequence >

         < xs:element name = ""PersonIdGuid"" type = ""t:GuidType"" minOccurs = ""0"" />
         < xs:element name = ""PersonId"" type = ""xs:int"" minOccurs = ""0"" />
         < xs:element name = ""FavoriteOrder"" type = ""xs:int"" minOccurs = ""0"" />
         < xs:element name = ""TrustLevel"" type = ""xs:int"" minOccurs = ""0"" />
         < xs:element name = ""RelevanceOrder1"" type = ""xs:string"" minOccurs = ""0"" />
         < xs:element name = ""RelevanceOrder2"" type = ""xs:string"" minOccurs = ""0"" />
         < xs:element name = ""AntiLinkInfo"" type = ""xs:string"" minOccurs = ""0"" />
         < xs:element name = ""ContactCategories"" type = ""t:ArrayOfStringsType"" minOccurs = ""0"" />
         < xs:element name = ""ContactHandles"" type = ""t:ArrayOfAbchPersonContactHandlesType"" minOccurs = ""0"" />
         < xs:element name = ""ExchangePersonIdGuid"" type = ""t:GuidType"" minOccurs = ""0"" />
       </ xs:sequence >
     </ xs:extension >
   </ xs:complexContent >
 </ xs:complexType >
");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R336002");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R336002
                // If the schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    336002,
                    @"[In Appendix C: Product Behavior] Implementation does support the AbchPersonItemType complex type which specifies a person. (Exchange 2016 and above follow this behavior.)");

                if (abchPersonItemType.AntiLinkInfo != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16005");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R16005
                    Site.CaptureRequirementIfIsInstanceOfType(
                        abchPersonItemType.AntiLinkInfo,
                        typeof(String),
                        16005,
                    @"[In t:AbchPersonItemType Complex Type] The type of child element AntiLinkInfo is xs:string ([XMLSCHEMA2]).");
                }

                if (abchPersonItemType.ContactCategories != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16011");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R16011
                    Site.CaptureRequirementIfIsInstanceOfType(
                        abchPersonItemType.ContactCategories,
                        typeof(String[]),
                        16011,
                    @"[In t:AbchPersonItemType Complex Type] The type of child element ContactCategories is t:ArrayOfStringsType ([MS-OXWSCDATA] section 2.2.4.13).");
                }

                if(abchPersonItemType.FavoriteOrderSpecified == true)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R16019");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R16019
                    Site.CaptureRequirementIfIsInstanceOfType(
                        abchPersonItemType.FavoriteOrder,
                        typeof(int),
                        16019,
                    @"[In t:AbchPersonItemType Complex Type] The type of child element FavoriteOrder is xs:int.");
                }
            }
        }
        #endregion

        #region Verify requirements related to CreateItem operation
        /// <summary>
        /// Capture CreateItemResponseType related requirements.
        /// </summary>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyCreateContactItem(bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R298");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R298
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                298,
                @"[In CreateItem] The following is the WSDL port type specification for the CreateItem operation:<wsdl:operation name=""CreateItem"">
                    <wsdl:input message=""tns:CreateItemSoapIn"" />
                    <wsdl:output message=""tns:CreateItemSoapOut"" />
                </wsdl:operation>");
        }
        #endregion

        #region Verify requirements related to ContactItemType complex types
        /// <summary>
        /// Capture ContactItemType Complex Type related requirements.
        /// </summary>
        /// <param name="contactItemType">A contact item from the response package of GetItem operation.</param>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyContactItemTypeComplexType(ContactItemType contactItemType, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema validation result should be true!");

            if (contactItemType != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R19");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R19
                // If the contactItemType element is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    19,
                    @"[In t: ContactItemType Complex Type] The type[ContactItemType] is defined as follow:
  < xs:complexType name = ""ContactItemType"" >
   < xs:complexContent >
     < xs:extension
       base = ""t: ItemType""
     >
       < xs:sequence >
         < xs:element name = ""FileAs""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""FileAsMapping""
           type = ""t: FileAsMappingType""
           minOccurs = ""0""
          />
         < xs:element name = ""DisplayName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""GivenName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Initials""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""MiddleName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Nickname""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""CompleteName""
           type = ""t: CompleteNameType""
           minOccurs = ""0""
          />
         < xs:element name = ""CompanyName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""EmailAddresses""
           type = ""t: EmailAddressDictionaryType""
           minOccurs = ""0""
          />
         < xs:element name = ""PhysicalAddresses""
           type = ""t: PhysicalAddressDictionaryType""
           minOccurs = ""0""
          />
         < xs:element name = ""PhoneNumbers""
           type = ""t: PhoneNumberDictionaryType""
           minOccurs = ""0""
          />
         < xs:element name = ""AssistantName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Birthday""
           type = ""xs: dateTime""
           minOccurs = ""0""
          />
	     <xs:element name=""BirthdayLocal""
              type = ""xs:dateTime""
              minOccurs = ""0""
          />
         < xs:element name = ""BusinessHomePage""
           type = ""xs: anyURI""
           minOccurs = ""0""
          />
         < xs:element name = ""Children""
           type = ""t: ArrayOfStringsType""
           minOccurs = ""0""
          />
         < xs:element name = ""Companies""
           type = ""t: ArrayOfStringsType""
           minOccurs = ""0""
          />
         < xs:element name = ""ContactSource""
           type = ""t: ContactSourceType""
           minOccurs = ""0""
          />
         < xs:element name = ""Department""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Generation""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""ImAddresses""
           type = ""t: ImAddressDictionaryType""
           minOccurs = ""0""
          />
         < xs:element name = ""JobTitle""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Manager""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Mileage""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""OfficeLocation""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""PostalAddressIndex""
           type = ""t: PhysicalAddressIndexType""
           minOccurs = ""0""
          />
         < xs:element name = ""Profession""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""SpouseName""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""Surname""
           type = ""xs: string""
           minOccurs = ""0""
          />
         < xs:element name = ""WeddingAnniversary""
           type = ""xs: dateTime""
           minOccurs = ""0""
          />
	     <xs:element name=""WeddingAnniversaryLocal""
              type = ""xs:dateTime""
              minOccurs = ""0""
          />
         < xs:element name = ""HasPicture""
           type = ""xs: boolean""
           minOccurs = ""0""
          />
         < xs:element name = ""PhoneticFullName""
           type = ""xs: string""
           minOccurs = ""0""
         />
         < xs:element name = ""PhoneticFirstName""
           type = ""xs: string""
           minOccurs = ""0""
         />
         < xs:element name = ""PhoneticLastName""
           type = ""xs: string""
           minOccurs = ""0""
         />
         < xs:element name = ""Alias""
           type = ""xs: string""
           minOccurs = ""0""
         />
         < xs:element name = ""Notes""
           type = ""xs: string""
           minOccurs = ""0""
         />
         < xs:element name = ""Photo""
           type = ""xs: base64Binary""
           minOccurs = ""0""
        />
        < xs:element name = ""UserSMIMECertificate""
          type = ""t: ArrayOfBinaryType""
          minOccurs = ""0""
        />
        < xs:element name = ""MSExchangeCertificate""
          type = ""t: ArrayOfBinaryType""
          minOccurs = ""0""
        />
        < xs:element name = ""DirectoryId""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""ManagerMailbox""
          type = ""t: SingleRecipientType""
          minOccurs = ""0""
        />
        < xs:element name = ""DirectReports""
          type = ""t: ArrayOfRecipientsType""
          minOccurs = ""0""
        />
        < xs:element name = ""AccountName""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""IsAutoUpdateDisabled""
          type = ""xs: boolean""
          minOccurs = ""0""
        />
        < xs:element name = ""IsMessengerEnabled""
          type = ""xs: boolean""
          minOccurs = ""0""
        />
        < xs:element name = ""Comment""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""ContactShortId""
          type = ""xs: int""
          minOccurs = ""0""
        />
        < xs:element name = ""ContactType""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""Gender""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""IsHidden""
          type = ""xs: boolean""
          minOccurs = ""0""
        />
        < xs:element name = ""ObjectId""
          type = ""xs: string""
          minOccurs = ""0""
        />
        < xs:element name = ""PassportId""
          type = ""xs: long""
          minOccurs = ""0""
        />
        < xs:element name = ""IsPrivate""
          type = ""xs: boolean""
          minOccurs = ""0""
        />
        < xs:element name = ""SourceId""
          type = ""xs: string""
          minOccurs = ""0""
        />
<xs:element name=""TrustLevel""
          type = ""xs:int""
          minOccurs = ""0""
        />
        < xs:element name = ""CreatedBy""
          type = ""xs:string""
          minOccurs = ""0""
        />
        < xs:element name = ""Urls""
          type = ""t:ContactUrlDictionaryType""
          minOccurs = ""0""
        />
        < xs:element name = ""AbchEmailAddresses""
           type = ""t: AbchEmailAddressDictionaryType""
           minOccurs = ""0""
          />
        < xs:element name = ""Cid""
          type = ""xs:long""
          minOccurs = ""0""
        />
        < xs:element name = ""SkypeAuthCertificate""
          type = ""xs:string""
          minOccurs = ""0""
        />
        < xs:element name = ""SkypeContext""
          type = ""xs:string""
          minOccurs = ""0""
        />
        < xs:element name = ""SkypeId""
          type = ""xs:string""
          minOccurs = ""0""
        />
        < xs:element name = ""SkypeRelationship""
          type = ""xs:string""
          minOccurs = ""0""
        />

            < xs:element name = ""YomiNickname""

              type = ""xs:string""

              minOccurs = ""0""
            />

            < xs:element name = ""XboxLiveTag""

              type = ""xs:string""

              minOccurs = ""0""
            />

            < xs:element name = ""InviteFree""

              type = ""xs:boolean""

              minOccurs = ""0""
            />

            < xs:element name = ""HidePresenceAndProfile""

              type = ""xs:boolean""

              minOccurs = ""0""
            />

            < xs:element name = ""IsPendingOutbound""

              type = ""xs:boolean""

              minOccurs = ""0""
            />

            < xs:element name = ""SupportGroupFeeds""

              type = ""xs:boolean""

              minOccurs = ""0""
            />

            < xs:element name = ""UserTileHash""

              type = ""xs:string""

              minOccurs = ""0""
            />

            < xs:element name = ""UnifiedInbox""

              type = ""xs:boolean""

              minOccurs = ""0""
            />

            < xs:element name = ""Mris""

              type = ""t:ArrayOfStringsType""

              minOccurs = ""0""
        />


            < xs:element name = ""Wlid""

              type = ""xs:string""

              minOccurs = ""0""
            />

            < xs:element name = ""AbchContactId""

              type = ""t:GuidType""

              minOccurs = ""0""
            />

            < xs:element name = ""NotInBirthdayCalendar""

               type = ""xs:boolean""

               minOccurs = ""0""
            />
        <
    xs:element name = ""ShellContactType""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""ImMri""

               type = ""xs:int""

               minOccurs = ""0""
            />

            < xs:element name = ""PresenceTrustLevel""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""OtherMri""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""ProfileLastChanged""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""MobileImEnabled""

               type = ""xs:boolean""

               minOccurs = ""0""
            />

            < xs:element name = ""DisplayNamePrefix""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""YomiGivenName""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""YomiSurname""

               type = ""xs:string""

               minOccurs = ""0""
            />

            < xs:element name = ""PersonalNotes""

               type = ""xs:string""

               minOccurs = ""0""
            />
            < xs:element name = ""PersonId""
              type = ""xs: int""
              minOccurs = ""0""
            />

          </ xs:sequence >

        </ xs:entension >

      </ xs:complexContent >

    </ xs:complexType >
");
            }

            if (Common.IsRequirementEnabled(1275004, this.Site) && contactItemType.AccountName != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334001");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334001
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.AccountName,
                    typeof(string),
                    334001,
                @"[In t:ContactItemType Complex Type] The type of element AccountName is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275004");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275004
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275004,
                    @"[In Appendix C: Product Behavior] Implementation does support AccountName element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275006, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334003");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334003
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.IsAutoUpdateDisabled,
                    typeof(Boolean),
                    334003,
                    @"[In t:ContactItemType Complex Type] The type of element IsAutoUpdateDisabled is xs:boolean.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275006");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275006
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275006,
                    @"[In Appendix C: Product Behavior] Implementation does support the IsAutoUpdateDisabled element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275008, this.Site) && contactItemType.Comment != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334007");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334007
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.Comment,
                    typeof(string),
                    334007,
                    @"[In t:ContactItemType Complex Type] The type of element Comment is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275008");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275008
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275008,
                    @"[In Appendix C: Product Behavior] Implementation does support the Comment element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275012, this.Site) && contactItemType.ContactType != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334007");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334011
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.ContactType,
                    typeof(string),
                    334011,
                    @"[In t:ContactItemType Complex Type] The type of element ContactType is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275012");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275012
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275012,
                    @"[In Appendix C: Product Behavior] Implementation does support the ContactType element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275014, this.Site) && contactItemType.Gender != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334013");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.Gender,
                    typeof(string),
                    334013,
                    @"[In t:ContactItemType Complex Type] The type of element Gender is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275014");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275014
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275014,
                    @"[In Appendix C: Product Behavior] Implementation does support the Gender element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275018, this.Site) && contactItemType.ObjectId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334017");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.ObjectId,
                    typeof(string),
                    334017,
                    @"[In t:ContactItemType Complex Type] The type of element ObjectId is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275018");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275018
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275018,
                    @"[In Appendix C: Product Behavior] Implementation does support the ObjectId element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275026, this.Site) && contactItemType.SourceId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334025");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.SourceId,
                    typeof(string),
                    334025,
                    @"[In t:ContactItemType Complex Type] The type of element SourceId is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275026");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275026
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275026,
                    @"[In Appendix C: Product Behavior] Implementation does support the SourceId element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275034, this.Site) && contactItemType.CidSpecified)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334033");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.Cid,
                    typeof(long),
                    334033,
                    @"[In t:ContactItemType Complex Type] The type of element Cid is xs:long.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275034");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275034
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275034,
                    @"[In Appendix C: Product Behavior] Implementation does support the Cid element. (Exchange 2016 and above follow this behavior.)");
            }
            if(contactItemType.PersonId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334021");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R334021
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.PersonId,
                    typeof(ItemIdType),
                    334021,
                    @"[In t:ContactItemType Complex Type] The type of element PersonId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");

            }
            if (Common.IsRequirementEnabled(1275036, this.Site) && contactItemType.SkypeAuthCertificate != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334035");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.SkypeAuthCertificate,
                    typeof(string),
                    334035,
                    @"[In t:ContactItemType Complex Type] The type of element SkypeAuthCertificate is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275036");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275036
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275036,
                    @"[In Appendix C: Product Behavior] Implementation does support the SkypeAuthCertificate element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275040, this.Site) && contactItemType.SkypeId != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334039");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.SkypeId,
                    typeof(string),
                    334039,
                    @"[In t:ContactItemType Complex Type] The type of element SkypeId is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275040");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275040
                // If schema is validated, the requirement can be validated.
                Site.CaptureRequirement(
                    1275040,
                    @"[In Appendix C: Product Behavior] Implementation does support the SkypeId element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275044, this.Site) && contactItemType.YomiNickname != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334043");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334043
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.YomiNickname,
                    typeof(string),
                    334043,
                @"[In t:ContactItemType Complex Type] The type of element YomiNickname is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275044");

                // Verify MS - OXWSCONT requirement: MS - OXWSCONT_R1275044
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275044,
                    @"[In Appendix C: Product Behavior] Implementation does support the YomiNickname element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275118, this.Site) && contactItemType.YomiGivenName != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334069");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334069
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.YomiGivenName,
                    typeof(string),
                    334069,
                @"[In t:ContactItemType Complex Type] The type of element YomiGivenName is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275118");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275118
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275118,
                    @"[In Appendix C: Product Behavior] Implementation does support the YomiGivenName  element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275120, this.Site) && contactItemType.YomiSurname != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334071");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334071
                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.YomiSurname,
                    typeof(string),
                    334071,
                    @"[In t:ContactItemType Complex Type] The type of element YomiSurname is xs:string.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275120");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275120
                // If schema is validated, the requirement can be validated.
                this.Site.CaptureRequirement(
                    1275120,
                    @"[In Appendix C: Product Behavior] Implementation does support the YomiSurname element. (Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1275116, this.Site) && contactItemType.DisplayNamePrefix != null)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334067");

                Site.CaptureRequirementIfIsInstanceOfType(
                    contactItemType.DisplayNamePrefix,
                    typeof(string),
                    334067,
                    @"[In t:ContactItemType Complex Type] The type of element DisplayNamePrefix is xs:string.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275116");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275116
                // If the DisplayNamePrefix element is specified and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    1275116,
                    @"[In Appendix C: Product Behavior] Implementation does support the DisplayNamePrefix  element. (Exchange 2016 and above follow this behavior.)");
            }

            if (contactItemType.FileAsMappingSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R128");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R128
                // If the FileAsMapping element is specified and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    128,
                    @"[In t:FileAsMappingType Simple Type] The type [FileAsMappingType] is defined as follow:
                        <xs:simpleType name=""FileAsMappingType"">
                         <xs:restriction
                          base=""xs:string""
                         >
                          <xs:enumeration
                           value=""None""
                           />
                          <xs:enumeration
                           value=""LastCommaFirst""
                           />
                          <xs:enumeration
                           value=""FirstSpaceLast""
                           />
                          <xs:enumeration
                           value=""Company""
                           />
                          <xs:enumeration
                           value=""LastCommaFirstCompany""
                           />
                          <xs:enumeration
                           value=""CompanyLastFirst""
                           />
                          <xs:enumeration
                           value=""LastFirst""
                           />
                          <xs:enumeration
                           value=""LastFirstCompany""
                           />
                          <xs:enumeration
                           value=""CompanyLastCommaFirst""
                           />
                          <xs:enumeration
                           value=""LastFirstSuffix""
                           />
                          <xs:enumeration
                           value=""LastSpaceFirstCompany""
                           />
                          <xs:enumeration
                           value=""CompanyLastSpaceFirst""
                           />
                          <xs:enumeration
                           value=""LastSpaceFirst""
                           />
                          <xs:enumeration
                           value=""DisplayName""
                           />
                          <xs:enumeration
                           value=""FirstName""
                           />
                          <xs:enumeration
                           value=""LastFirstMiddleSuffix""
                           />
                          <xs:enumeration
                           value=""LastName""
                           />
                          <xs:enumeration
                           value=""Empty""
                           />
                         </xs:restriction>
                        </xs:simpleType>");

                if (Common.IsRequirementEnabled(1275096, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275096");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275096
                    // If the FileAsMapping element is specified and schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275096,
                        @"[In Appendix C: Product Behavior] Implementation does support the DisplayName attribute. (Exchange 2010 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(1275098, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275098");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275098
                    // If the FileAsMapping element is specified and schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275098,
                        @"[In Appendix C: Product Behavior] Implementation does support the FirstName attribute. (Exchange 2010 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(1275100, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275100");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275100
                    // If the FileAsMapping element is specified and schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275100,
                        @"[In Appendix C: Product Behavior] Implementation does support the LastFirstMiddleSuffix attribute. (Exchange 2010 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(1275102, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275102");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275102
                    // If the FileAsMapping element is specified and schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275102,
                        @"[In Appendix C: Product Behavior] Implementation does support the LastName attribute. (Exchange 2010 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled(1275104, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275104");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275104
                    // If the FileAsMapping element is specified and the schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275104,
                        @"[In Appendix C: Product Behavior] Implementation does support the Empty attribute. (Exchange 2010 and above follow this behavior.)");
                }
            }

            if (Common.IsRequirementEnabled(1275032, this.Site) && contactItemType.Urls != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224011");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224011
                // If the Urls element is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    224011,
                    @"[In t:ContactUrlDictionaryType Complex Type] The type [ContactUrlDictionaryType] is defined as follow:
<xs:complexType name=""ContactUrlDictionaryType"" >
   < xs:sequence >
     < xs:element name = ""Url"" type = ""t:ContactUrlDictionaryEntryType"" maxOccurs = ""unbounded"" />
   </ xs:sequence >
 </ xs:complexType > ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275032");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275032
                // If the Urls element is not null and the schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    1275032,
                    @"[In Appendix C: Product Behavior] Implementation does support the Urls element. (Exchange 2016 and above follow this behavior.)");

                if (Common.IsRequirementEnabled(1275084, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R334031");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R334031
                    // If the Urls element is not null and the schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        334031,
                        @"[In t:ContactItemType Complex Type] The type of element Urls is t:ContactUrlDictionaryType (section 3.1.4.1.1.9).");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275084");

                    // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275084
                    // If the Urls element is not null and the schema is validated,
                    // the requirement can be validated.
                    Site.CaptureRequirement(
                        1275084,
                        @"[In Appendix C: Product Behavior] Implementation does support the ContactUrlDictionaryType complex type. (Exchange 2016 and above follow this behavior.).");
                }

                for (int i = 0; i < contactItemType.Urls.Length; i++)
                {
                    if (contactItemType.Urls[i] != null)
                    {
                        if (Common.IsRequirementEnabled(1275082, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275082");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275082
                            // If the entry of Urls is not null and schema is validated,
                            // the requirement can be validated.
                            Site.CaptureRequirement(
                                1275082,
                                @"[In Appendix C: Product Behavior] Implementation does support the ContactUrlDictionaryEntryType complex type. (Exchange 2016 and above follow this behavior.)");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R224002");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R224002
                            // If the entry of Urls is not null and schema is validated,
                            // the requirement can be validated.
                            Site.CaptureRequirement(
                                224002,
                                @"[In t:ContactUrlDictionaryEntryType Complex Type]  The type [ContactUrlDictionaryEntryType] is defined as follow:
 <xs:complexType name=""ContactUrlDictionaryEntryType"" >
   < xs:sequence >
     < xs:element name = ""Type"" type = ""t:ContactUrlKeyType"" minOccurs = ""1"" />
     < xs:element name = ""Address"" type = ""xs:string"" minOccurs = ""0"" />
     < xs:element name = ""Name"" type = ""xs:string"" minOccurs = ""0"" />
   </ xs:sequence >
 </ xs:complexType >
");
                        }

                        if (Common.IsRequirementEnabled(1275094, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275094");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275094
                            // The Type element is a required element of ContactUrlDictionaryEntryType, if the entry as ContactUrlDictionaryEntryType is not null,
                            // and the schema is validated, this requirement can be validated.                            
                            Site.CaptureRequirement(
                                1275094,
                                @"[In Appendix C: Product Behavior] Implementation does support the ContactUrlKeyType simple type. (Exchange 2016 and above follow this behavior.)");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R120009");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R120009
                            // The Type element is a required element of ContactUrlDictionaryEntryType, if the entry as ContactUrlDictionaryEntryType is not null,
                            // and the schema is validated, this requirement can be validated.                            
                            Site.CaptureRequirement(
                            120009,
                            @"[In t:ContactUrlKeyType Simple Type] The type [ContactUrlKeyType] is defined as follow:
<xs:simpleType name=""ContactUrlKeyType"" >
< xs:restriction base = ""xs:string"" >
    < xs:enumeration value = ""Personal"" />
    < xs:enumeration value = ""Business"" />
    < xs:enumeration value = ""Attachment"" />
    < xs:enumeration value = ""EbcDisplayDefinition"" />
    < xs:enumeration value = ""EbcFinalImage"" />
    < xs:enumeration value = ""EbcLogo"" />
    < xs:enumeration value = ""Feed"" />
    < xs:enumeration value = ""Image"" />
    < xs:enumeration value = ""InternalMarker"" />
    < xs:enumeration value = ""Other"" />
</ xs:restriction >
</ xs:simpleType >
");
                        }
                        else
                        {
                            Site.Assert.Fail("The entry of PhysicalAddresses should not be null!");
                        }
                    }
                }
            }

            if (contactItemType.Companies != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1081");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1081
                // If the Companies element is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1081,
                    @"[In t:ArrayOfStringsType Complex Type] The type [ArrayOfStringsType] is defined as follow:
                        <xs:complexType name=""ArrayOfStringsType"">
                          <xs:sequence>
                            <xs:element name=""String"" type=""xs:string"" minOccurs=""0"" maxOccurs=""unbounded""/>
                          </xs:sequence>
                        </xs:complexType>");
            }

            if (contactItemType.PostalAddressIndexSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R178");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R178
                // If the PostalAddressIndex element is specified and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    178,
                    @"[t:PhysicalAddressIndexType Simple Type] The type [PhysicalAddressIndexType] is defined as follow:<xs:simpleType name=""PhysicalAddressIndexType"">
                         <xs:restriction
                          base=""xs:string""
                         >
                          <xs:enumeration
                           value=""None""
                           />
                          <xs:enumeration
                           value=""Business""
                           />
                          <xs:enumeration
                           value=""Home""
                           />
                          <xs:enumeration
                           value=""Other""
                           />
                         </xs:restriction>
                        </xs:simpleType>");
            }

            if (contactItemType.CompleteName != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R192");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R192
                // If the CompleteName element is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    192,
                    @"[In t:CompleteNameType Complex Type] The type [CompleteNameType] is defined as follow:<xs:complexType name=""CompleteNameType"">
                         <xs:sequence>
                          <xs:element name=""Title""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""FirstName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""MiddleName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""LastName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""Suffix""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""Initials""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""FullName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""Nickname""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""YomiFirstName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                          <xs:element name=""YomiLastName""
                           type=""xs:string""
                           minOccurs=""0""
                           />
                         </xs:sequence>
                        </xs:complexType>");
            }

            if (contactItemType.EmailAddresses != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R236");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R236
                // If the EmailAddresses is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    236,
                    @"[In t:EmailAddressDictionaryType Complex Type] The type [EmailAddressDictionaryType] is defined as follow:<xs:complexType name=""EmailAddressDictionaryType"">
                         <xs:sequence>
                          <xs:element name=""Entry""
                           type=""t:EmailAddressDictionaryEntryType""
                           maxOccurs=""unbounded""
                           />
                         </xs:sequence>
                        </xs:complexType>");

                for (int i = 0; i < contactItemType.EmailAddresses.Length; i++)
                {
                    if (contactItemType.EmailAddresses[i] != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R226");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R226
                        // If the entry of EmailAddresses is not null and schema is validated,
                        // the requirement can be validated.
                        Site.CaptureRequirement(
                            226,
                            @"[In t: EmailAddressDictionaryEntryType Complex Type] The type[EmailAddressDictionaryEntryType] is defined as follow:
 < xs:complexType name = ""EmailAddressDictionaryEntryType"" >
 < xs:simpleContent >
  < xs:extension
   base = ""xs: string""
  >
   < xs:attribute name = ""Key""
    type = ""t: EmailAddressKeyType""
    use = ""required""
    />
   < xs:attribute name = ""Name""
    type = ""xs: string""
    use = ""optional""
    />
   < xs:attribute name = ""RoutingType""
    type = ""xs: string""
    use = ""optional""
    />
   < xs:attribute name = ""MailboxType""
    type = ""t: MailboxTypeType""
    use = ""optional""
    />
  </ xs:extension >
 </ xs:simpleContent >
</ xs:complexType > ");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R122");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R122
                        // The Key element is a required element of EmailAddressDictionaryEntryType, if the entry as EmailAddressDictionaryEntryType is not null,
                        // and the schema is validated, this requirement can be validated.
                        Site.CaptureRequirement(
                            122,
                            @"[In t:EmailAddressKeyType Simple Type] The type [EmailAddressKeyType] is defined as follow:
                                <xs:simpleType name=""EmailAddressKeyType"">
                                 <xs:restriction
                                  base=""xs:string""
                                 >
                                  <xs:enumeration
                                   value=""EmailAddress1""
                                   />
                                  <xs:enumeration
                                   value=""EmailAddress2""
                                   />
                                  <xs:enumeration
                                   value=""EmailAddress3""
                                   />
                                 </xs:restriction>
                                </xs:simpleType>");

                        if (Common.IsRequirementEnabled(1275086, this.Site) && contactItemType.EmailAddresses[i].Name !=null)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275086");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275086
                            // If the Name is not null and the schema is validated, this requirement can be validated.
                            Site.CaptureRequirement(
                                1275086,
                                @"[In Appendix C: Product Behavior] Implementation does support the Name attribute. (Exchange 2010 and above follow this behavior.)");
                        }

                        if (Common.IsRequirementEnabled(1275088, this.Site) && contactItemType.EmailAddresses[i].RoutingType != null)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275088");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275088
                            // If the Name is not null and the schema is validated, this requirement can be validated.
                            Site.CaptureRequirement(
                                1275088,
                                @"[In Appendix C: Product Behavior] Implementation does support the Name attribute. (Exchange 2010 and above follow this behavior.)");
                        }

                        if (Common.IsRequirementEnabled(1275090, this.Site) && contactItemType.EmailAddresses[i].MailboxTypeSpecified == true)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275090");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275090
                            // If the Name is not null and the schema is validated, this requirement can be validated.
                            Site.CaptureRequirement(
                                1275090,
                                @"[In Appendix C: Product Behavior] Implementation does support the Name attribute. (Exchange 2010 and above follow this behavior.)");
                        }
                    }
                    else
                    {
                        Site.Assert.Fail("The entry of EmailAddresses should not be null!");
                    }
                }
            }

            if (contactItemType.ImAddresses != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R244");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R244
                // If the ImAddresses is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    244,
                    @"[In t:ImAddressDictionaryType Complex Type] The type [ImAddressDictionaryType] is defined as follow:<xs:complexType name=""ImAddressDictionaryType"">
                         <xs:sequence>
                          <xs:element name=""Entry""
                           type=""t:ImAddressDictionaryEntryType""
                           maxOccurs=""unbounded""
                           />
                         </xs:sequence>
                        </xs:complexType>");

                for (int i = 0; i < contactItemType.ImAddresses.Length; i++)
                {
                    if (contactItemType.ImAddresses[i] != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R240");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R240
                        // If the entry of ImAddresses is not null and schema is validated,
                        // the requirement can be validated.
                        Site.CaptureRequirement(
                            240,
                            @"[In t:ImAddressDictionaryEntryType Complex Type] The type [ImAddressDictionaryEntryType] is defined as follow:<xs:complexType name=""ImAddressDictionaryEntryType"">
                                 <xs:simpleContent>
                                  <xs:extension
                                   base=""xs:string""
                                  >
                                   <xs:attribute name=""key""
                                    type=""t:ImAddressKeyType""
                                    use=""required""
                                    />
                                  </xs:extension>
                                 </xs:simpleContent>
                                </xs:complexType>");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R149");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R149
                        // The Key element is a required element of ImAddressDictionaryEntryType, if the entry as ImAddressDictionaryEntryType is not null,
                        // and the schema is validated, this requirement can be validated.
                        Site.CaptureRequirement(
                            149,
                            @"[In t:ImAddressKeyType Simple Type] The type [ImAddressKeyType] is defined as follow:
                                <xs:simpleType name=""ImAddressKeyType"">
                                 <xs:restriction
                                  base=""xs:string""
                                 >
                                  <xs:enumeration
                                   value=""ImAddress1""
                                   />
                                  <xs:enumeration
                                   value=""ImAddress2""
                                   />
                                  <xs:enumeration
                                   value=""ImAddress3""
                                   />
                                 </xs:restriction>
                                </xs:simpleType>");
                    }
                    else
                    {
                        Site.Assert.Fail("The entry of ImAddresses should not be null!");
                    }
                }
            }

            if (contactItemType.PhoneNumbers != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R252");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R252
                // If the PhoneNumbers is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    252,
                    @"[In t:PhoneNumberDictionaryType Complex Type] The type [PhoneNumberDictionaryType] is defined as follow:<xs:complexType name=""PhoneNumberDictionaryType"">
                     <xs:sequence>
                      <xs:element name=""Entry""
                       type=""t:PhoneNumberDictionaryEntryType""
                       maxOccurs=""unbounded""
                       />
                     </xs:sequence>
                    </xs:complexType>");

                for (int i = 0; i < contactItemType.PhoneNumbers.Length; i++)
                {
                    if (contactItemType.PhoneNumbers[i] != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R248");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R248
                        // If the entry of PhoneNumbers is not null and schema is validated,
                        // the requirement can be validated.
                        Site.CaptureRequirement(
                            248,
                            @"[In t:PhoneNumberDictionaryEntryType Complex Type] The type [PhoneNumberDictionaryEntryType] is defined as follow:<xs:complexType name=""PhoneNumberDictionaryEntryType"">
                                <xs:simpleContent>
                                <xs:extension
                                base=""xs:string""
                                >
                                <xs:attribute name=""Key""
                                type=""t:PhoneNumberKeyType""
                                use=""required""
                                />
                                </xs:extension>
                                </xs:simpleContent>
                            </xs:complexType>");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R155");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R155
                        // The Key element is a required element of PhoneNumberDictionaryEntryType, if the entry as PhoneNumberDictionaryEntryType is not null,
                        // and the schema is validated, this requirement can be validated.
                        Site.CaptureRequirement(
                            155,
                            @"[In t: PhoneNumberKeyType Simple Type] The type[PhoneNumberKeyType] is defined as follow:
< xs:simpleType name = ""PhoneNumberKeyType"" >
 < xs:restriction
  base = ""xs: string""
 >
  < xs:enumeration
   value = ""AssistantPhone""
   />
  < xs:enumeration
   value = ""BusinessFax""
   />
  < xs:enumeration
   value = ""BusinessPhone""
   />
  < xs:enumeration
   value = ""BusinessPhone2""
   />
  < xs:enumeration
   value = ""Callback""
   />
  < xs:enumeration
   value = ""CarPhone""
   />
  < xs:enumeration
   value = ""CompanyMainPhone""
   />
  < xs:enumeration
   value = ""HomeFax""
   />
  < xs:enumeration
   value = ""HomePhone""
   />
  < xs:enumeration
   value = ""HomePhone2""
   />
  < xs:enumeration
   value = ""Isdn""
   />
  < xs:enumeration
   value = ""MobilePhone""
   />
  < xs:enumeration
   value = ""OtherFax""
   />
  < xs:enumeration
   value = ""OtherTelephone""
   />
  < xs:enumeration
   value = ""Pager""
   />
  < xs:enumeration
   value = ""PrimaryPhone""
   />
  < xs:enumeration
   value = ""RadioPhone""
   />
  < xs:enumeration
   value = ""Telex""
   />
  < xs:enumeration
   value = ""TtyTddPhone""
   />
   < xs:enumeration
       value = ""BusinessMobile"" />
     < xs:enumeration
       value = ""IPPhone"" />
     < xs:enumeration
       value = ""Mms"" />
     < xs:enumeration
       value = ""Msn"" />

 </ xs:restriction >
</ xs:simpleType > ");

                        if (Common.IsRequirementEnabled(1275106, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275106");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275106
                            // If BusinessMobile is not null and the schema is verifed, this requirement can be validated.
                            Site.CaptureRequirementIfIsNotNull(
                                PhoneNumberKeyType.BusinessMobile,
                                1275106,
                                @"[In Appendix C: Product Behavior] Implementation does support  the IPPhone value. (Exchange 2016 and above follow this behavior.)");
                        }

                            if (Common.IsRequirementEnabled(1275108, this.Site))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275108");

                            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275108
                            // If IPPhone is not null and the schema is verifed, this requirement can be validated.
                            Site.CaptureRequirementIfIsNotNull(
                                PhoneNumberKeyType.IPPhone,
                                1275108,
                                @"[In Appendix C: Product Behavior] Implementation does support  the IPPhone value. (Exchange 2016 and above follow this behavior.)");
                        }
                    }
                    else
                    {
                        Site.Assert.Fail("The entry of PhoneNumbers should not be null!");
                    }
                }
            }

            if (contactItemType.PhysicalAddresses != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R270");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R270
                // If the PhysicalAddresses is not null and schema is validated,
                // the requirement can be validated.
                Site.CaptureRequirement(
                    270,
                    @"[In t:PhysicalAddressDictionaryType Complex Type] The type [PhysicalAddressDictionaryType] is defined as follow:<xs:complexType name=""PhysicalAddressDictionaryType"">
                         <xs:sequence>
                          <xs:element name=""entry""
                           type=""t:PhysicalAddressDictionaryEntryType""
                           maxOccurs=""unbounded""
                           />
                         </xs:sequence>
                        </xs:complexType>");

                for (int i = 0; i < contactItemType.PhysicalAddresses.Length; i++)
                {
                    if (contactItemType.PhysicalAddresses[i] != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R256");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R256
                        // If the entry of PhysicalAddresses is not null and schema is validated,
                        // the requirement can be validated.
                        Site.CaptureRequirement(
                            256,
                            @"[In t:PhysicalAddressDictionaryEntryType Complex Type] The type [PhysicalAddressDictionaryEntryType] is defined as follow:<xs:complexType name=""PhysicalAddressDictionaryEntryType"">
                                 <xs:sequence>
                                  <xs:element name=""Street""
                                   type=""xs:string""
                                   minOccurs=""0""
                                   />
                                  <xs:element name=""City""
                                   type=""xs:string""
                                   minOccurs=""0""
                                   />
                                  <xs:element name=""State""
                                   type=""xs:string""
                                   minOccurs=""0""
                                   />
                                  <xs:element name=""CountryOrRegion""
                                   type=""xs:string""
                                   minOccurs=""0""
                                   />
                                  <xs:element name=""PostalCode""
                                   type=""xs:string""
                                   minOccurs=""0""
                                   />
                                 </xs:sequence>
                                 <xs:attribute name=""Key""
                                  type=""t:PhysicalAddressKeyType""
                                  use=""required""
                                  />
                                </xs:complexType>");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R185");

                        // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R185
                        // The Key element is a required element of PhysicalAddressDictionaryEntryType, if the entry as PhysicalAddressDictionaryEntryType is not null,
                        // and the schema is validated, this requirement can be validated.
                        Site.CaptureRequirement(
                            185,
                            @"[In t:PhysicalAddressKeyType Simple Type] The type [PhysicalAddressKeyType] is defined as follow:<xs:simpleType name=""PhysicalAddressKeyType"">
                                 <xs:restriction
                                  base=""xs:string""
                                 >
                                  <xs:enumeration
                                   value=""Business""
                                   />
                                  <xs:enumeration
                                   value=""Home""
                                   />
                                  <xs:enumeration
                                   value=""Other""
                                   />
                                 </xs:restriction>
                                </xs:simpleType>");
                    }
                    else
                    {
                        Site.Assert.Fail("The entry of PhysicalAddresses should not be null!");
                    }
                }
            }
        }
        #endregion

        #region Verify BaseResponseMessageType Structure
        /// <summary>
        /// Verify the BaseResponseMessageType structure.
        /// </summary>
        /// <param name="baseResponseMessage">A BaseResponseMessageType instance.</param>
        private void VerifyBaseResponseMessageType(BaseResponseMessageType baseResponseMessage)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1081001");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1081001
            Site.CaptureRequirementIfIsNotNull(
                baseResponseMessage,
                "MS-OXWSCDATA",
                1081001,
                @"[In m:BaseResponseMessageType Complex Type] The type [BaseResponseMessageType] is defined as follow:
                <xs:complexType name=""BaseResponseMessageType"">
                  <xs:sequence>
                    <xs:element name=""ResponseMessages""
                      type=""m:ArrayOfResponseMessagesType""
                     />
                  </xs:sequence>
                </xs:complexType>");
        }
        #endregion

        #region Verify ResponseMessageType Structure
        /// <summary>
        /// Verify the ResponseMessageType structure.
        /// </summary>
        /// <param name="responseMessage">A ResponseMessageType instance.</param>
        private void VerifyResponseMessageType(ResponseMessageType responseMessage)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R114701");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R114701
            Site.CaptureRequirementIfIsNotNull(
                responseMessage,
                "MS-OXWSCDATA",
                114701,
                @"[In m:ResponseMessageType Complex Type] The type [ResponseMessageType] is defined as follow:
                    <xs:complexType name=""ResponseMessageType"">
                      <xs:sequence
                        minOccurs=""0""
                      >
                        <xs:element name=""MessageText""
                          type=""xs:string""
                          minOccurs=""0""
                         />
                        <xs:element name=""ResponseCode""
                          type=""m:ResponseCodeType""
                          minOccurs=""0""
                         />
                        <xs:element name=""DescriptiveLinkKey""
                          type=""xs:int""
                          minOccurs=""0""
                         />
                        <xs:element name=""MessageXml""
                          minOccurs=""0""
                        >
                          <xs:complexType>
                            <xs:sequence>
                              <xs:any
                                process_contents=""lax""
                                minOccurs=""0""
                                maxOccurs=""unbounded""
                               />
                            </xs:sequence>
                            <xs:attribute name=""ResponseClass""
                              type=""t:ResponseClassType""
                              use=""required""
                             />
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>");
        }

        #region Verify ServerVersionInfo Structure
        /// <summary>
        /// Verify the ServerVersionInfo structure.
        /// </summary>
        /// <param name="serverVersionInfo">A ServerVersionInfo instance.</param>
        /// <param name="isSchemaValidated">Indicate whether the schema is verified.</param>
        private void VerifyServerVersionInfo(ServerVersionInfo serverVersionInfo, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368004");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368004
            Site.CaptureRequirementIfIsNotNull(
                serverVersionInfo,
                "MS-OXWSCDATA",
                368004,
                @"[In t:ServerVersionInfo Element] <xs:element name=""t:ServerVersionInfo"">
                      <xs:complexType>
                        <xs:attribute name=""MajorVersion""
                          type=""xs:int""
                          use=""optional""
                         />
                        <xs:attribute name=""MinorVersion""
                          type=""xs:int""
                          use=""optional""
                         />
                        <xs:attribute name=""MajorBuildNumber""
                          type=""xs:int""
                          use=""optional""
                         />
                        <xs:attribute name=""MinorBuildNumber""
                          type=""xs:int""
                          use=""optional""
                         />
                        <xs:attribute name=""Version""
                          type=""xs:string""
                          use=""optional""
                         />
                      </xs:complexType>
                    </xs:element>");

            if (serverVersionInfo.MajorVersionSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368005");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368005
                // If MajorVersion element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368005,
                    "[In t:ServerVersionInfo Element] The type of the attribute MajorVersion is xs:int ([XMLSCHEMA2])");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368006
                // If MinorVersion element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    serverVersionInfo.MajorVersion,
                    "MS-OXWSCDATA",
                    368006,
                    @"[In t:ServerVersionInfo Element] MajorVersion attribute: Specifies the server's major version number.");
            }

            if (serverVersionInfo.MinorVersionSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368007");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368007
                // If MinorVersion element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368007,
                    "[In t:ServerVersionInfo Element] The type of the attribute MinorVersion is xs:int ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368008");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368008
                // If MinorVersion element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    serverVersionInfo.MinorVersion,
                    "MS-OXWSCDATA",
                    368008,
                    "[In t:ServerVersionInfo Element] MinorVersion attribute: Specifies the server's minor version number.");
            }

            if (serverVersionInfo.MajorBuildNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368009");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368009
                // If MajorBuildNumber element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368009,
                    "[In t:ServerVersionInfo Element] The type of the attribute MajorBuildNumber is xs:int");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368010");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368010
                // If MajorBuildNumber element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    serverVersionInfo.MajorBuildNumber,
                    "MS-OXWSCDATA",
                    368010,
                    "[In t:ServerVersionInfo Element] MajorBuildNumber attribute: Specifies the server's major build number.");
            }

            if (serverVersionInfo.MinorBuildNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368011");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368011
                // If MinorBuildNumber element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368011,
                    "[In t:ServerVersionInfo Element] The type of the attribute MinorBuildNumber is xs:int");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368012");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368012
                // If MinorBuildNumber element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    serverVersionInfo.MinorBuildNumber,
                    "MS-OXWSCDATA",
                    368012,
                    "[In t:ServerVersionInfo Element] MinorBuildNumber attribute: Specifies the server's minor build number.");
            }

            if (serverVersionInfo.Version != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368013");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368013
                // If Version element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368013,
                    "[In t:ServerVersionInfo Element] The type of the attribute Version is xs:string ([XMLSCHEMA2])");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R368014");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R368014
                // If MS-OXWSCDATA_r368013 is verifed successfully, this requirement can be validated directly
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    368014,
                    "[In t:ServerVersionInfo Element] Version attribute: specifies the server's version number including the major version number, minor version number, major build number, and minor build number in that order.");
            }
        }
        #endregion
        #endregion

        #region Verify SetUserPhotoMessageResponseType Structure
        /// <summary>
        /// Capture SetuserPhotoResponseMessageType related requirements.
        /// </summary>
        /// <param name="getUserPhotoMessageResponse">Specified SetUserPhotoResponseMessageType instance.</param>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifySetUserPhotoResponseMessageType(SetUserPhotoResponseMessageType setUserPhotoMessageResponse, bool isSchemaValidated)
        {
            // Verify the base type ResponseMessageType related requirements.
            this.VerifyResponseMessageType(setUserPhotoMessageResponse as ResponseMessageType);

            // If the schema validation and the above base type verification are successful, then MS-OXWSCONT_R302126 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302126");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R302126
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                302126,
                @"[In SetUserPhotoResponseMessageType] The following is the SetUserPhotoResponseMessageType complex type specification. 
   <xs:complexType name=""SetUserPhotoResponseMessageType"" >
     < xs:complexContent >
       < xs:extension base = ""m:ResponseMessageType"" />
     </ xs:complexContent >
 <  / xs:complexType >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275114");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275114
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1275114,
                @"[In Appendix C: Product Behavior] Implementation does support the SetUserPhoto operation. (Exchange 2016 and above follow this behavior.)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302125");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302125
            this.Site.CaptureRequirementIfIsNotNull(
                setUserPhotoMessageResponse,
                302125,
                @"[In SetUserPhotoResponseMessageType] This type extends the ResponseMessageType complex type, as specified by [MS-OXWSCDATA] section 2.2.4.67.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302078");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302078
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                302078,
                @"[In SetUserPhoto] The following is the WSDL port type specification of the SetUserPhoto WSDL operation.
   <wsdl:operation name=""SetUserPhoto"" >
     < wsdl:input message = ""tns:SetUserPhotoSoapIn"" />
     < wsdl:output message = ""tns:SetUserPhotoSoapOut"" />
   </ wsdl:operation >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302079");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302079
            Site.CaptureRequirementIfIsNotNull(
                setUserPhotoMessageResponse,
                302079,
                @"[In SetUserPhoto] The following is the WSDL binding specification of the SetUserPhoto WSDL operation.
   <wsdl:operation name=""SetUserPhoto"" >
     < soap:operation soapAction = ""http://schemas.microsoft.com/exchange/services/2006/messages/SetUserPhoto"" />
     < wsdl:input >
       < soap:header message = ""tns:SetUserPhotoSoapIn"" part = ""RequestVersion"" use = ""literal"" />
       < soap:body parts = ""request"" use = ""literal"" />
     </ wsdl:input >
     < wsdl:output >
       < soap:body parts = ""SetUserPhotoResult"" use = ""literal"" />
       < soap:header message = ""tns:SetUserPhotoSoapOut"" part = ""ServerVersion"" use = ""literal"" />
     </ wsdl:output >
   </ wsdl:operation >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302094");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302094
            Site.CaptureRequirementIfIsNotNull(
                setUserPhotoMessageResponse,
                302094,
                @"[In SetUserPhotoSoapOut] The following is the SetUserPhotoSoapOut WSDL message specification.
   <wsdl:message name=""SetUserPhotoSoapOut"" >
       < wsdl:part name = ""SetUserPhotoResult"" element = ""tns:SetUserPhotoResponse"" />
       < wsdl:part name = ""ServerVersion"" element = ""t:ServerVersionInfo"" />
     </ wsdl:message >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302097");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302097
            Site.CaptureRequirementIfIsNotNull(
                setUserPhotoMessageResponse,
                302097,
                @"[In SetUserPhotoSoapOut] The element of the part SetUserPhotoResult is tns:SetUserPhotoResponse (section 3.1.4.8.2.2)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302098");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302098
            // According to the schema, getUserPhotoMessageResponse is the SOAP body of a response message returned by server, this requirement can be verified directly.
            Site.CaptureRequirement(
                302098,
                @"[In SetUserPhotoSoapOut] SetUserPhotoResult part: Specifies the SOAP body of the response to a SetUserPhoto WSDL operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302099");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302099
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                302099,
                @"[In SetUserPhotoSoapOut] The element of the part ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302100");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302100
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, this requirement can be verified directly.
            Site.CaptureRequirement(
                302100,
                @"[In SetUserPhotoSoapOut] ServerVersion part: Specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302107");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302107
            Site.CaptureRequirementIfIsNotNull(
                setUserPhotoMessageResponse,
                302107,
                @"[In SetUserPhotoResponse] The following is the SetUserPhotoResponse element specification.
 <xs:element name=""SetUserPhotoResponse"" type =""m: SetUserPhotoResponseMessageType"" />
");
        }
        #endregion

        #region Verify GetUserPhotoMessageResponseType Structure
        /// <summary>
        /// Capture GetuserPhotoResponseMessageType related requirements.
        /// </summary>
        /// <param name="getUserPhotoMessageResponse">Specified GetUserPhotoResponseMessageType instance.</param>
        /// <param name="isSchemaValidated">A boolean value indicates the schema validation result. True means the response conforms with the schema, false means not.</param>
        private void VerifyGetUserPhotoResponseMessageType(GetUserPhotoResponseMessageType getUserPhotoMessageResponse, bool isSchemaValidated)
        {
            // Verify the base type ResponseMessageType related requirements.
            this.VerifyResponseMessageType(getUserPhotoMessageResponse as ResponseMessageType);

            // If the schema validation and the above base type verification are successful, then MS-OXWSCONT_R302053, MS-OXWSCONT_R302055, MS-OXWSCONT_R302057 and MS-OXWSCONT_R1275124 can be captured.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302053");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R302053
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                302053,
                @"[In GetUserPhotoResponseMessageType] The following is the GetUserPhotoResponseMessageType complex type specification. 
   <xs:complexType name=""GetUserPhotoResponseMessageType"" >
     < xs:complexContent >
       < xs:extension base = ""m:ResponseMessageType"" >
         < xs:sequence >
           < xs:element name = ""HasChanged"" type = ""xs:boolean""
                 minOccurs = ""1"" maxOccurs = ""1"" />
           < xs:element name = ""PictureData"" type = ""xs:base64Binary""
                 minOccurs = ""0"" maxOccurs = ""1"" />
           < xs:element name = ""ContentType"" type = ""xs:string""
                 minOccurs = ""0"" maxOccurs = ""1"" />
         </ xs:sequence >
       </ xs:extension >
     </ xs:complexContent >
   </ xs:complexType >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302052");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302052
            Site.CaptureRequirementIfIsNotNull(
                getUserPhotoMessageResponse,
                302052,
                @"[In GetUserPhotoResponseMessageType] This type [GetUserPhotoResponseMessageType] extends the ResponseMessageType complex type, as specified in [MS-OXWSCDATA] section 2.2.4.67.");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302055");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302055
            Site.CaptureRequirementIfIsInstanceOfType(
                getUserPhotoMessageResponse.HasChanged,
                typeof(Boolean),
                302055,
                @"[In GetUserPhotoResponseMessageType] The type of the element HasChanged is xs:boolean ([XMLSCHEMA2])");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302057");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302057
            Site.CaptureRequirementIfIsInstanceOfType(
                getUserPhotoMessageResponse.PictureData,
                typeof(byte[]),
                302057,
                @"[In GetUserPhotoResponseMessageType] The type of the element PictureData is xs:base64Binary ([XMLSCHEMA2])");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302155");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302155
            Site.CaptureRequirementIfIsInstanceOfType(
                getUserPhotoMessageResponse.ContentType,
                typeof(String),
                302155,
                @"[In GetUserPhotoResponseMessageType] The type of the element ContentType is xs:string ([XMLSCHEMA2])");

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R1275124");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1275124
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                1275124,
                @"[In Appendix C: Product Behavior] Implementation does support  the ContentType element. (Exchange 2013 and above follow this behavior.)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R302002");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302002
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                302002,
                @"[In GetUserPhoto] The following is the WSDL port type specification of the GetUserPhoto WSDL operation.
 <wsdl:operation name=""GetUserPhoto"" >
       < wsdl:input message = ""tns:GetUserPhotoSoapIn"" />
       < wsdl:output message = ""tns:GetUserPhotoSoapOut"" />
     </ wsdl:operation >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302003");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302003
            Site.CaptureRequirementIfIsNotNull(
                getUserPhotoMessageResponse,
                302003,
                @"[In GetUserPhoto] The following is the WSDL binding specification of the GetUserPhoto WSDL operation.
 <wsdl:operation name=""GetUserPhoto"" >
       < soap:operation soapAction =
     ""http://schemas.microsoft.com/exchange/services/2006/messages/GetUserPhoto"" />
       < wsdl:input >
         < soap:body parts = ""request"" use = ""literal"" />
         < soap:header message = ""tns:GetUserPhotoSoapIn""
               part = ""RequestVersion"" use = ""literal"" />
       </ wsdl:input >
       < wsdl:output >


         < soap:header message = ""tns:GetUserPhotoSoapOut""
               part = ""ServerVersion"" use = ""literal"" />
        < soap:body parts = ""GetUserPhotoResult"" use = ""literal"" />
       </ wsdl:output >
     </ wsdl:operation > ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302017");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302017
            Site.CaptureRequirementIfIsNotNull(
                getUserPhotoMessageResponse,
                302017,
                @"[In GetUserPhotoSoapOut] The following is the GetUserPhotoSoapOut WSDL message specification.
   <wsdl:message name=""GetUserPhotoSoapOut"" >
     < wsdl:part name = ""GetUserPhotoResult"" element = ""tns:GetUserPhotoResponse"" />
     < wsdl:part name = ""ServerVersion"" element = ""t:ServerVersionInfo"" />
   </ wsdl:message >
");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302020");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302020
            Site.CaptureRequirementIfIsNotNull(
                getUserPhotoMessageResponse,
                302020,
                @"[In GetUserPhotoSoapOut] The element of the part GetUserPhotoResult is tns:GetUserPhotoResponse (section 3.1.4.7.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302021");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302021
            // According to the schema, getUserPhotoMessageResponse is the SOAP body of a response message returned by server, this requirement can be verified directly.
            Site.CaptureRequirement(
                302021,
                @"[In GetUserPhotoSoapOut] GetUserPhotoResult part: Specifies the SOAP body of the response to a GetUserPhoto WSDL operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302022");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302022
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                302022,
                @"[In GetUserPhotoSoapOut] The element of the part ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302023");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302023
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, this requirement can be verified directly.
            Site.CaptureRequirement(
                302023,
                @"[In GetUserPhotoSoapOut] ServerVersion part: Specifies a SOAP header that identifies the server version for the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R302034");

            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R302034
            Site.CaptureRequirementIfIsNotNull(
                getUserPhotoMessageResponse,
                302034,
                @"[In GetUserPhotoResponse] The following is the GetUserPhotoResponse element specification.
 <xs:element name=""GetUserPhotoResponse""
         type = ""m:GetUserPhotoResponseMessageType""
 xmlns: xs = ""http://www.w3.org/2001/XMLSchema"" />
");
        }
        #endregion

        #region Verify requirements related to SOAP version
        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R1
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                1,
                @"[In Transport] The SOAP version supported is SOAP 1.1, as specified in [SOAP1.1].");
        }
        #endregion

        #region Verify transport related requirements.
        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);
            if (Common.IsRequirementEnabled(335001, this.Site) && transport == TransportProtocol.HTTPS)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R335001");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R335001
                // When test suite running on HTTPS, if there are no exceptions or error messages returned from server, this requirement will be captured.
                Site.CaptureRequirement(
                    335001,
                    @"[In Appendix B: Product Behavior] Implementation does use secure communications via HTTPS, as defined in [RFC2818]. (Exchange 2007 and above follow this behavior.)");
            }

            if (transport == TransportProtocol.HTTP)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCONT_R101");

                // Verify MS-OXWSCONT requirement: MS-OXWSCONT_R101
                // When test suite running on HTTP, if there are no exceptions or error messages returned from server, this requirement will be captured.
                Site.CaptureRequirement(
                    101,
                    @"[In Transport] The protocol MUST support SOAP over HTTP, as specified in [RFC2616]. ");
            }
        }
        #endregion
    }
}