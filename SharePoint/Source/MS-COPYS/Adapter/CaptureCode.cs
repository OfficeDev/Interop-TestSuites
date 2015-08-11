namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Text;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The partial class of adapter of MS-COPYS, it implements the adapter capture code.
    /// </summary>
    public partial class MS_COPYSAdapter
    {   
        /// <summary>
        /// A method used to verify GetItem operation related captures.
        /// </summary>
        /// <param name="responseOfGetItem">A parameter represents the response of a GetItem operation.</param>
        private void VerifyGetItemOperationCapture(GetItemResponse responseOfGetItem)
        {
            if (null == responseOfGetItem)
            {
                throw new ArgumentNullException("responseOfGetItem");
            }

            this.VerifyTransPortAndSOAPCapture();

            // If schema validation is successful, the request and response schema definition related requirements: MS-COPYS_R169, MS-COPYS_R171, MS-COPYS_R183, MS-COPYS_R184, MS-COPYS_R191, MS-COPYS_R192 and MS-COPYS_R195 can be directly captured.
            this.Site.CaptureRequirement(
                169,
                @"[In GetItem] [The GetItem schema is:]: <wsdl:operation name=""GetItem"">
    <wsdl:input message=""tns:GetItemSoapIn"" />
    <wsdl:output message=""tns:GetItemSoapOut"" />
</wsdl:operation>");

            // Verified requirement: MS-COPYS_R171
            this.Site.CaptureRequirement(
                171,
                @"[In GetItem] [The protocol client sends a GetItemSoapIn request message (section 3.1.4.1.1.1) and] the protocol server responds with a GetItemSoapOut response message (section 3.1.4.1.1.2) as follows:");

            // Verified requirement: MS-COPYS_R183
            this.Site.CaptureRequirement(
                183,
                @"[In GetItemSoapOut] The SOAP action value of the message is defined as:
http://schemas.microsoft.com/sharepoint/soap/GetItem");

            // Verified requirement: MS-COPYS_R184
            this.Site.CaptureRequirement(
                184,
                @"[In GetItemSoapOut] The SOAP body contains a GetItemResponse element (section 3.1.4.1.2.2).");

            // Verified requirement: MS-COPYS_R191
            this.Site.CaptureRequirement(
                191,
                @"[In GetItemResponse] It contains content and metadata for the requested file.");

            // Verified requirement: MS-COPYS_R192
            this.Site.CaptureRequirement(
                192,
                @"[In GetItemResponse] [The schema of the GetItemResponse is defined as:] <s:element name=""GetItemResponse"">
              <s:complexType>
                <s:sequence>
                  <s:element minOccurs=""1"" maxOccurs=""1""
                    name=""GetItemResult"" type=""s:unsignedInt"" />
                  <s:element minOccurs=""0"" maxOccurs=""1""
                    name=""Fields"" type=""tns:FieldInformationCollection"" />
                  <s:element minOccurs=""0"" maxOccurs=""1""
                    name=""Stream"" type=""s:base64Binary"" />
                </s:sequence>
              </s:complexType>
            </s:element>");

            this.Site.CaptureRequirementIfIsTrue(
                responseOfGetItem.GetItemResult.Equals(0),
                195,
                @"[In GetItemResponse] [GetItemResult] The protocol server MUST set this value to zero.");

          // Validate the field information collection and field information items.
            this.VerifyFieldInfromationCollection(responseOfGetItem.Fields);
        }

        /// <summary>
        /// A method used to verify CopyIntoItemsLocal operation related captures. 
        /// </summary>
        /// <param name="response">A parameter represents the response of CopyIntoItemsLocal operation.</param>
        /// <param name="destinationUrlsInRequest">A parameter represents the destination URL items which is passed in the request.</param>
        private void VerifyCopyIntoItemsLocalOperationCapture(CopyIntoItemsLocalResponse response, string[] destinationUrlsInRequest)
        {
            if (null == response)
            {
                throw new ArgumentNullException("response");
            }

            this.VerifyTransPortAndSOAPCapture();

            // If schema validation is successful, the request and response schema definition related requirements: MS-COPYS_R268, MS-COPYS_R295, MS-COPYS_R296, MS-COPYS_R297, MS-COPYS_R306, MS-COPYS_R307, MS-COPYS_R266 and MS-COPYS_R309 can be directly captured. 
            this.Site.CaptureRequirement(
                268,
                @"[In CopyIntoItemsLocal] [The protocol client sends a CopyIntoItemsLocalSoapIn request message (section 3.1.4.3.1.1) and] the protocol server responds with a CopyIntoItemsLocalSoapOut response message (section 3.1.4.3.1.2) as follows:");

            // Verified requirement: MS-COPYS_R266
            this.Site.CaptureRequirement(
                266,
                @"[In CopyIntoItemsLocal] [The schema of the CopyIntoItemsLocal is defined as: ] <wsdl:operation name=""CopyIntoItemsLocal"">
    <wsdl:input message=""tns:CopyIntoItemsLocalSoapIn"" />
    <wsdl:output message=""tns:CopyIntoItemsLocalSoapOut"" />
</wsdl:operation>");

            // Verified requirement: MS-COPYS_R295
            this.Site.CaptureRequirement(
                295,
                @"[In CopyIntoItemsLocalSoapOut] The CopyIntoItemsLocalSoapOut message is the response WSDL message for a CopyIntoItemsLocal WSDL operation (section 3.1.4.3).");

            // Verified requirement: MS-COPYS_R296
            this.Site.CaptureRequirement(
                296,
                @"[In CopyIntoItemsLocalSoapOut] The SOAP action value of the message is defined as:
http://schemas.microsoft.com/sharepoint/soap/CopyIntoItemsLocal");

            // Verified requirement: MS-COPYS_R297
            this.Site.CaptureRequirement(
                297,
                @"[In CopyIntoItemsLocalSoapOut] The SOAP body contains a CopyIntoItemsLocalResponse element (section 3.1.4.3.2.2).");

            // Verified requirement: MS-COPYS_R306
            this.Site.CaptureRequirement(
                306,
                @"[In CopyIntoItemsLocalResponse] [The schema of the CopyIntoItemsLocalResponse is defined as: ] <s:element name=""CopyIntoItemsLocalResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""CopyIntoItemsLocalResult"" type=""s:unsignedInt"" />
      <s:element minOccurs=""0"" maxOccurs=""1"" name=""Results"" type=""tns:CopyResultCollection"" />
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verified requirement: MS-COPYS_R307
            this.Site.CaptureRequirement(
                307,
                @"[In CopyIntoItemsLocalResponse] CopyIntoItemsLocalResult: The result of the CopyIntoItemsLocal operation.");

            // Verified requirement: MS-COPYS_R309
            this.Site.CaptureRequirement(
                309,
                @"[In CopyIntoItemsLocalResponse] [CopyIntoItemsLocalResult] The protocol server MUST set this value to zero (""0"").");

            // Verified requirement: MS-COPYS_R304
            this.Site.CaptureRequirement(
                304,
                @"[In CopyIntoItemsLocal] [DestinationUrls] The CopyIntoItemsLocalResponse element specifies a protocol server response for the CopyIntoItemsLocal operation (section 3.1.4.3).");

            // Verify Copy result collection and all copy result items in the collection
            this.VerifyCopyResultCollection(response.Results, destinationUrlsInRequest);
        }

        /// <summary>
        /// A method used to verify CopyIntoItems operation related captures. 
        /// </summary>
        /// <param name="response">A parameter represents the response of CopyIntoItems operation.</param>
        /// <param name="destinationUrlsInRequest">A parameter represents the destination URL items which is passed in the request.</param>
        private void VerifyCopyIntoItemsOperationCapture(CopyIntoItemsResponse response, string[] destinationUrlsInRequest)
        {
            if (null == response)
            {
                throw new ArgumentNullException("response");
            }

            this.VerifyTransPortAndSOAPCapture();

            // If schema validation is successful, the request and response schema definition related requirements: MS-COPYS_R204, MS-COPYS_R239, MS-COPYS_R240, MS-COPYS_R241, MS-COPYS_R255, MS-COPYS_R258 can be directly captured. 
            this.Site.CaptureRequirement(
                204,
                @"[In CopyIntoItems] [The protocol client sends a CopyIntoItemsoapIn request message (section 3.1.4.2.1.1) and] the protocol server responds with a CopyIntoItemsoapOut response message (section 3.1.4.2.1.2 ) as follows.");

            // Verified requirement: MS-COPYS_R239
            this.Site.CaptureRequirement(
                239,
                @"[In CopyIntoItemsSoapOut] The CopyIntoItemsSoapOut message is the response WSDL message for a CopyIntoItems WSDL operation (section 3.1.4.2).");

            // Verified requirement: MS-COPYS_R240
            this.Site.CaptureRequirement(
                240,
                @"[In CopyIntoItemsSoapOut] The SOAP action value of the message is defined as:
http://schemas.microsoft.com/sharepoint/soap/CopyIntoItems");

            // Verified requirement: MS-COPYS_R241
            this.Site.CaptureRequirement(
                241,
                @"[In CopyIntoItemsSoapOut] The SOAP body contains a CopyIntoItemsResponse element (section 3.1.4.2.2.2).");

            // Verified requirement: MS-COPYS_R202
            this.Site.CaptureRequirement(
                202,
                @"[In CopyIntoItems] [The schema of the operation CopyIntoItems is defined as:] <wsdl:operation name=""CopyIntoItems"">
                <wsdl:input message=""tns:CopyIntoItemsoapIn"" />
                <wsdl:output message=""tns:CopyIntoItemsoapOut"" />
            </wsdl:operation>");

            // Verified requirement: MS-COPYS_R255
            this.Site.CaptureRequirement(
                255,
                @"[In CopyIntoItemsResponse] [The schema of the CopyIntoItemsResponse is defined as: ]  <s:element name=""CopyIntoItemsResponse"">
  <s:complexType>
    <s:sequence>
      <s:element minOccurs=""1"" maxOccurs=""1"" name=""CopyIntoItemsResult"" type=""s:unsignedInt"" />
      <s:element minOccurs=""0"" maxOccurs=""1"" name=""Results"" type=""tns:CopyResultCollection"" />
    </s:sequence>
  </s:complexType>
</s:element>");

            this.Site.CaptureRequirementIfIsTrue(
                response.CopyIntoItemsResult.Equals(0),
                258,
                @"[In CopyIntoItemsResponse] [CopyIntoItemsResult] The protocol server MUST set this value to zero (""0"").");

            // Verify Copy result collection and all copy result items in the collection
            this.VerifyCopyResultCollection(response.Results, destinationUrlsInRequest);
        }

        /// <summary>
        /// A method used to verify the transPort and SOAP related captures.
        /// </summary>
        private void VerifyTransPortAndSOAPCapture()
        {   
            // If the test suite send a request and then receive a response successfully, then capture SOAP version related requirements
            switch (TestSuiteManageHelper.CurrentSoapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        // Verified requirement: MS-COPYS_R6
                        this.Site.CaptureRequirement(
                                                   6,
                                                   @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1] section 4.");
                        break;
                    }

                case SoapVersion.SOAP12:
                    {
                        // Verified requirement: MS-COPYS_R7
                        this.Site.CaptureRequirement(
                                                   7,
                                                   @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.2/1] section 5.");
                        break;
                    }
            }

            // If the test suite send a request and then receive a response successfully, then capture transport version related requirements
            switch (TestSuiteManageHelper.CurrentTransportType)
            {
                case TransportProtocol.HTTP:
                    {
                        // Verified requirement: MS-COPYS_R3
                        this.Site.CaptureRequirement(
                                                 3,
                                                 @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
                        break;
                    }

                case TransportProtocol.HTTPS:
                    {   
                        if (Common.IsRequirementEnabled(5, this.Site))
                        {
                            // Verified requirement: MS-COPYS_R5
                            this.Site.CaptureRequirement(
                                                 5,
                                                 @"[In Appendix B: Product Behavior] Implementation does additionally support SOAP over HTTPS to help secure communication with protocol clients. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                        }
                       
                        break;
                    }
            }
        }

        /// <summary>
        /// A method used to verify field information collection and all field items in this collection.
        /// </summary>
        /// <param name="fieldInfromationCollection">A parameter represents the  field information collection.</param>
        private void VerifyFieldInfromationCollection(FieldInformation[] fieldInfromationCollection)
        {
            if (null == fieldInfromationCollection || 0 == fieldInfromationCollection.Length)
            {
                return;
            }

            // If the schema validation is successful, the requirements: MS-COPYS_R76 and MS-COPYS_R77 can be directly captured.
            this.Site.CaptureRequirement(
                76,
                @"[In FieldInformationCollection] [The schema of complex type FieldInformationCollection is defined as:] <s:complexType name=""FieldInformationCollection"">
  <s:sequence>
    <s:element name=""FieldInformation"" type=""tns:FieldInformation"" minOccurs=""0"" maxOccurs=""unbounded""/>
  </s:sequence>
</s:complexType>");

            // Verified requirement: MS-COPYS_R77
            this.Site.CaptureRequirement(
                77,
                @"[In FieldInformationCollection] FieldInformation: A single metadata field for a file, as defined in section 2.2.4.4.");

            // If the schema validation is successful, all required properties in each fieldinformation item in the FieldInformationCollection are match the schema definition, the requirements: MS-COPYS_R62, MS-COPYS_R64, MS-COPYS_R65, MS-COPYS_R67 and MS-COPYS_R69 can be captured.
            this.Site.CaptureRequirement(
                62,
                @"[In FieldInformation] [The schema of complex type FieldInformation is defined as:] <s:complexType name=""FieldInformation"">
                  <s:attribute name=""Type"" type=""tns:FieldType"" use=""required""/>
                  <s:attribute name=""DisplayName"" type=""s:string"" use=""required""/>
                  <s:attribute name=""InternalName"" type=""s:string"" use=""required""/>
                  <s:attribute name=""Id"" type=""s1:guid"" use=""required""/>
                  <s:attribute name=""Value"" type=""s:string""/>
                </s:complexType>");

            string[] types = { "Invalid", "Integer", "Text", "Note", "DateTime", "Counter", "Choice", "Lookup", "Boolean", "Number", "Currency", "URL", "Computed", "Threading", "Guid", "MultiChoice", "GridChoice", "Calculated", "File", "Attachments", "User", "Recurrence", "CrossProjectLink", "ModStat", "AllDayEvent", "Error" };
            
            bool isVerifyR64 = false;
            foreach (string type in types)
            {
                if (fieldInfromationCollection[0].Type.ToString() == type)
                {
                    isVerifyR64 = true;
                    break;
                }
            }

            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR64,
                64,
                @"[In FieldInformation] [Type] The value MUST contain one of the values [Invalid, Integer, Text, Note, DateTime, Counter, Choice, Lookup, Boolean, Number, Currency, URL, Computed, Threading, Guid, MultiChoice, GridChoice, Calculated, File, Attachments, User, Recurrence, CrossProjectLink, ModStat, AllDayEvent, Error] defined in section 2.2.5.2.");

            // The simple type FieldType
            this.Site.CaptureRequirementIfIsTrue(
                isVerifyR64,
                127,
                @"[In FieldType] Attributes that use FieldType MUST use one of the following values: Invalid, Integer, Text, Note, DateTime, Counter, Choice, Lookup, Boolean, Number, Currency, URL, Computed, Threading, Guid, MultiChoice, GridChoice, Calculated, File, Attachments, User, Recurrence, CrossProjectLink, ModStat, AllDayEvent, Error.");

            // Verified requirement: MS-COPYS_R65
            this.Site.CaptureRequirement(
                65,
                @"[In FieldInformation] DisplayName: The user-readable name of the field.");

            // Verified requirement: MS-COPYS_R67
            this.Site.CaptureRequirement(
                67,
                @"[In FieldInformation] InternalName: The internal name that identifies the metadata field for a file in the source location.");

            // Verified requirement: MS-COPYS_R69
            this.Site.CaptureRequirement(
                69,
                @"[In FieldInformation] Id: The GUID that identifies the metadata field for a file in the source location.");

            // If the schema validation is successful, all Fieldinformation item ins the FieldInformationCollection are match the schema definition, the requirements: MS-COPYS_R113, MS-COPYS_R127 and MS-COPYS_R129 can be directly captured.
            this.Site.CaptureRequirement(
                113,
                @"[In FieldType] [The schema of simple  type FieldType is defined as:] <s:simpleType name=""FieldType"">
  <s:restriction base=""s:string"">
    <s:enumeration value=""Invalid""/>
    <s:enumeration value=""Integer""/>
    <s:enumeration value=""Text""/>
    <s:enumeration value=""Note""/>
    <s:enumeration value=""DateTime""/>
    <s:enumeration value=""Counter""/>
    <s:enumeration value=""Choice""/>
    <s:enumeration value=""Lookup""/>
    <s:enumeration value=""Boolean""/>
    <s:enumeration value=""Number""/>
    <s:enumeration value=""Currency""/>
    <s:enumeration value=""URL""/>
    <s:enumeration value=""Computed""/>
    <s:enumeration value=""Threading""/>
    <s:enumeration value=""Guid""/>
    <s:enumeration value=""MultiChoice""/>
    <s:enumeration value=""GridChoice""/>
    <s:enumeration value=""Calculated""/>
    <s:enumeration value=""File""/>
    <s:enumeration value=""Attachments""/>
    <s:enumeration value=""User""/>
    <s:enumeration value=""Recurrence""/>
    <s:enumeration value=""CrossProjectLink""/>
    <s:enumeration value=""ModStat""/>
    <s:enumeration value=""AllDayEvent""/>
    <s:enumeration value=""Error""/>
  </s:restriction>
</s:simpleType>");

            // Verified requirement: MS-COPYS_R129
            this.Site.CaptureRequirement(
                129,
                @"[In guid] [The schema of simple  type guid is defined as:] <s:simpleType name=""guid"">
  <s:restriction base=""s:string"">
    <s:pattern value=""[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"" />
  </s:restriction>
</s:simpleType>");

            // Verify each field item.
            foreach (FieldInformation fieldItem in fieldInfromationCollection)
            {
                this.VerifyFieldInformation(fieldItem);
            }
        }

        /// <summary>
        /// A method used to verify field information related requirements.
        /// </summary>
        /// <param name="fieldInformationItem">A parameter represents the field information.</param>
        private void VerifyFieldInformation(FieldInformation fieldInformationItem)
        {
            if (null == fieldInformationItem)
            {
                throw new ArgumentNullException("fieldInformationItem");
            }
            
            if (!string.IsNullOrEmpty(fieldInformationItem.Value))
            {
                // If the value attribute presents, then capture R70
                this.Site.CaptureRequirement(
                                            70,
                                            "[In FieldInformation] Value: The value of the field.");
            }

            // If the display name is non-empty and fewer than 256 characters. then capture R66.
            // The unicode encoding format can present all possible chars, all string are unicode string.
            this.Site.Log.Add(
                        LogEntryKind.Debug,
                        @"The DisplayName[{0}] length is [{1}]",
                        string.IsNullOrEmpty(fieldInformationItem.DisplayName) ? "Null/Empty" : fieldInformationItem.DisplayName,
                        string.IsNullOrEmpty(fieldInformationItem.DisplayName) ? "NUll/Zero" : fieldInformationItem.DisplayName.Length.ToString());
             
            this.Site.CaptureRequirementIfIsTrue(
                                                !string.IsNullOrEmpty(fieldInformationItem.DisplayName) && fieldInformationItem.DisplayName.Length < 256,
                                                66,
                                                "[In FieldInformation] [DisplayName] This value MUST be a non-empty Unicode string that is fewer than 256 characters.");

            // If the internal name is a ASCII string ,capture R68
            this.Site.CaptureRequirementIfIsTrue(
                                              this.VerifyInternalNameOfField(fieldInformationItem.InternalName),
                                              68,
                                              "[In FieldInformation] [InternalName] The value MUST be a non-empty ASCII string that does not contain spaces and is fewer than 256 characters.");
        }
 
        /// <summary>
        /// A method used to verify the internal name of a field whether is ASCII string, its length is fewer than 256 and does not contain any space.
        /// </summary>
        /// <param name="internalName">A parameter represents the internal name of a field.</param>
        /// <returns>Return 'true' indicating the internal name is an ASCII string, its length is fewer than 256 and does not contain any space.</returns>
        private bool VerifyInternalNameOfField(string internalName)
        {
            if (string.IsNullOrEmpty(internalName))
            {
                throw new ArgumentNullException("internalName");
            }
 
            if (internalName.Length >= 256)
            {
                this.Site.Log.Add(
                                 LogEntryKind.Debug,
                                 "This length of internal Name[{0}] should be fewer than 256. Actual:[{1}]",
                                 internalName,
                                 internalName.Length);
                return false;
            }

            if (internalName.IndexOf(" ") >= 0)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                "This internal Name[{0}] should not contain any space.",
                                internalName);
                return false;
            }

            bool isValidationError = false;

            #region validate the internalName
            char[] charsOfStringValue = internalName.ToCharArray();

            foreach (char charItem in charsOfStringValue)
            {
                char[] validatedChars = new char[] { charItem };
                byte[] encodedBytesOfUniCode = Encoding.Unicode.GetBytes(validatedChars);
                byte[] encodedBytesOfAscii = Encoding.ASCII.GetBytes(validatedChars);

                // If it is not "Little-Endian order", change it to Little-Endian order
                if (!BitConverter.IsLittleEndian)
                {
                    Array.Reverse(encodedBytesOfUniCode);
                }

                // The high byte is the second byte item in the UNICODE byte array whose value should be equal to 0, if the char is ASCII table scope char.
                if (encodedBytesOfUniCode[1] != 0)
                {
                    this.Site.Log.Add(
                                      LogEntryKind.Debug,
                                      @"This char[{0}] in internalName[{1}] is not a valid ASCII char. Its ""most significant byte"" should be equal to 0. actual bytes:[low{2}- most{3}].",
                                      charItem,
                                      internalName,
                                      encodedBytesOfUniCode[0],
                                      encodedBytesOfUniCode[1]);
                    isValidationError = true;
                    break;
                }

                // The low byte is the first byte item in the UNICODE byte array whose value should be equal to ASCII byte.
                if (encodedBytesOfUniCode[0] != encodedBytesOfAscii[0])
                {
                    this.Site.Log.Add(
                                    LogEntryKind.Debug,
                                    @"This char[{0}] in internalName[{1}] is not a valid ASCII char. Its ""lowest significant byte"" should be equal to ASCII byte. actual bytes:[low{2}- ASSIC{3}].",
                                    charItem,
                                    internalName,
                                    encodedBytesOfUniCode[1],
                                    encodedBytesOfAscii[0]);
                    isValidationError = true;
                    break;
                }
            }

            #endregion validate the internalName

            return !isValidationError;
        }

        /// <summary>
        /// A method used to verify the CopyResultCollection type.
        /// </summary>
        /// <param name="copyResultCollection">A parameter represents the CopyResultCollection type instance which will be validated.</param>
        /// <param name="destinationUrlsInRequest">A parameter represents the destination URL items which is used to test destination URL mappings.</param>
        private void VerifyCopyResultCollection(CopyResult[] copyResultCollection, string[] destinationUrlsInRequest)
        {
            if (null == copyResultCollection || 0 == copyResultCollection.Length)
            {
                return;
            }

            // if the CopyResult is not null, then capture R59
            this.Site.CaptureRequirement(
                                        59,
                                        @"[In CopyResultCollection] CopyResult: Specifies the status of the copy operation for a single destination location, as defined in section 2.2.4.2.");
            
            // if the schema validation is successful, then capture R58
            this.Site.CaptureRequirement(
                                       58,
                                       @"[In CopyResultCollection] [The schema of complex type CopyResultCollection is defined as:] <s:complexType name=""CopyResultCollection"">
  <s:sequence>
    <s:element minOccurs=""0"" maxOccurs=""unbounded"" name=""CopyResult"" nillable=""true"" type=""tns:CopyResult"" />
  </s:sequence>
</s:complexType>");

            // if the schema validation is successful, all CopyResult items in the CopyResult collection should match the schema definition of CopyResult type.
            this.Site.CaptureRequirement(
                                        46,
                                        @"[In CopyResult] [The schema of complex type CopyResult is defined as:] <s:complexType name=""CopyResult"">
  <s:attribute name=""ErrorCode"" type=""tns:CopyErrorCode"" use=""required""/>
  <s:attribute name=""ErrorMessage"" type=""s:string""/>
  <s:attribute name=""DestinationUrl"" type=""s:string"" use=""required""/>
</s:complexType>");

            // if the schema validation is successful, all CopyResult items in the CopyResult collection should match the schema definition of CopyResult type.
            this.Site.CaptureRequirement(
                                        47,
                                        "[In CopyResult] ErrorCode: The success or failure status code of the operation for a specific DestinationUrl, as defined in section 2.2.5.1.");

            // if the schema validation is successful, all CopyResult items in the CopyResult collection should match the schema definition of CopyResult type.
            this.Site.CaptureRequirement(
                                        55,
                                        "[In CopyResult] DestinationUrl: The destination location for which the CopyResult element specifies the result of the operation.");

            // if the schema validation is successful, all CopyResult items in the CopyResult collection should match the schema definition of CopyResult type, the requirements: MS-COPYS_R84 and MS-COPYS_R85 can be directly captured.
            this.Site.CaptureRequirement(
                84,
                @"[In CopyErrorCode] [The schema of simple  type CopyErrorCode is defined as:] <s:simpleType name=""CopyErrorCode"">
  <s:restriction base=""s:string"">
    <s:enumeration value=""Success""/>
    <s:enumeration value=""DestinationInvalid""/>
    <s:enumeration value=""DestinationMWS""/>
    <s:enumeration value=""SourceInvalid""/>
    <s:enumeration value=""DestinationCheckedOut""/>
    <s:enumeration value=""InvalidUrl""/>
    <s:enumeration value=""Unknown""/>
  </s:restriction>
</s:simpleType>");

            // Verified requirement: MS-COPYS_R85
            this.Site.CaptureRequirement(
                85,
                @"[In CopyErrorCode] The following table lists the allowed values of the CopyErrorCode simple type[Success, DestinationInvalid, DestinationMWS, SourceInvalid, DestinationCheckedOut, InvalidUrl, Unknown].");
            
            #region validate each CopyResult item.
            // Validate the destinationUrl property in each CopyResult item whether match each item in DestinationUrlCollection instance in the request.
            if (null == destinationUrlsInRequest || 0 == destinationUrlsInRequest.Length)
            {
                throw new ArgumentNullException("detinationUrlsInRequest");
            }

            if (copyResultCollection.Length != destinationUrlsInRequest.Length)
            {
                this.Site.Assert.Fail(
                         "The number of URL items in DestinationUrlCollection instance should equal to the number of copyResult items in CopyResultCollection instance. Urls in request:[{0}], copyResult items[{1}]",
                         destinationUrlsInRequest.Length,
                         copyResultCollection.Length);
            }

            StringBuilder errorMsgOfValidateDestinationUrl = new StringBuilder();
            for (int itemIndex = 0; itemIndex < destinationUrlsInRequest.Length; itemIndex++)
            {
                if (!string.Equals(destinationUrlsInRequest[itemIndex], copyResultCollection[itemIndex].DestinationUrl, StringComparison.OrdinalIgnoreCase))
                {
                    // record information for un-equal destination Urls.
                    string errorMsg = string.Format(
                                                   @"Not matched item--Index:[{0}] URL in request:[{1}], URL in CopyResult:[{2}]\r\n",
                                                   itemIndex,
                                                   string.IsNullOrEmpty(destinationUrlsInRequest[itemIndex]) ? "None/Empty" : destinationUrlsInRequest[itemIndex],
                                                   string.IsNullOrEmpty(copyResultCollection[itemIndex].DestinationUrl) ? "None/Empty" : copyResultCollection[itemIndex].DestinationUrl);
                    errorMsgOfValidateDestinationUrl.AppendLine(errorMsg);
                }

                // If the "ErrorMessage" attribute has of current CopyResult item, then capture R48 for CopyResult item
                if (!string.IsNullOrEmpty(copyResultCollection[itemIndex].ErrorMessage))
                {
                    // Verified requirement: MS-COPYS_R48
                    this.Site.CaptureRequirement(
                                             48,
                                             "[In CopyResult] ErrorMessage: The user-readable message that explains the failure specified by the ErrorCode attribute.");
                }
            }

            if (errorMsgOfValidateDestinationUrl.Length != 0)
            {
                this.Site.Log.Add(
                               LogEntryKind.Debug,
                               @"Not all the destination URLs are matched between the DestinationUrl property of CopyResult item and the URL item in DestinationUrlCollection instance in the request:\r\n{0}",
                               errorMsgOfValidateDestinationUrl.ToString());
            }

            #endregion validate each CopyResult item.

            // If all the URLs values are matched between CopyResult collection and DestinationUrlCollection, then capture R57.
            this.Site.CaptureRequirementIfAreEqual(
                                               0,
                                               errorMsgOfValidateDestinationUrl.Length,
                                               57,
                                               "[In CopyResultCollection] This collection MUST contain exactly one entry for each IRI in the DestinationUrlCollection complex type (section 2.2.4.1).");
        }

        /// <summary>
        /// A method used to verify the detail element definition of a soap exception. 
        /// </summary>
        /// <param name="soapEx">A parameter represents the soap exception instance which should contain a "detail" element.</param>
        private void VerifySoapExceptionDetailCapture(SoapException soapEx)
        {
            // Validate the detail element schema definition.
            SoapFaultDetailSchemaHelper.ValidateSoapFaultDetail(soapEx, this.Site);

            // If there are no any schema issues, then capture the R22, R24
            this.Site.CaptureRequirement(
                                        22,
                                        "[In SOAP Fault Message] In a SOAP fault response, the detail element contains application-specific error information.");

            // Verified requirement: MS-COPYS_R24
            this.Site.CaptureRequirement(
                                        24,
                                        @"[In SOAP Fault Message] The following schema specifies the structure of the detail element in the SOAP fault response that is used by this protocol: <s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap"">
                                         <s:complexType name=""SOAPFaultDetails"">
                                            <s:sequence>
                                               <s:element name=""errorstring"" type=""s:string""/>
                                               <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
                                            </s:sequence>
                                         </s:complexType>
                                      </s:schema>");
        }
    }
}