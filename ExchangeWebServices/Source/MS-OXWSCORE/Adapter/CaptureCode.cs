namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using System.Text.RegularExpressions;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSCORE.
    /// </summary>
    public partial class MS_OXWSCOREAdapter
    {
        #region Verify soap version related requirements.
        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R3");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R3
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                3,
                @"[In Transport] The SOAP version supported is SOAP 1.1.");
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

            if (Common.IsRequirementEnabled(1006, this.Site) && transport == TransportProtocol.HTTPS)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1006");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1006
                // When test suite running on HTTPS, if there are no exceptions or error messages returned from server, this requirement will be captured.
                Site.CaptureRequirement(
                    1006,
                    @"[In Appendix C: Product Behavior] Implementation does use secure communications via HTTPS, as defined in [RFC2818]. (Exchange 2007 and above follow this behavior.)");
            }

            if (transport == TransportProtocol.HTTP)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2001");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2001
                // When test suite running on HTTP, if there are no exceptions or error messages returned from server, this requirement will be captured.
                Site.CaptureRequirement(
                    2001,
                    @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }
        }
        #endregion

        #region Verify Section 1: Protocol Details
        #region Verify CopyItemResponseTypes Structure
        /// <summary>
        /// Verify the CopyItemResponseTypes structure.
        /// </summary>
        /// <param name="copyItemResponse">A CopyItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyCopyItemResponse(CopyItemResponseType copyItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R222");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R222
            Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                222,
                @"[In CopyItem Operation] The following is the WSDL port type specification for the CopyItem operation: 
                    <wsdl:operation name=""CopyItem"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
                         <wsdl:input message=""tns:CopyItemSoapIn"" />
                         <wsdl:output message=""tns:CopyItemSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R223");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R223
            Site.CaptureRequirementIfIsNotNull(
                 copyItemResponse,
                223,
                @"[In CopyItem Operation] The following is the WSDL binding specification for the CopyItem operation: 
                    <wsdl:operation name=""CopyItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CopyItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""CopyItemResult"" use=""literal""/>
                          <soap:header message=""tns:CopyItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R237");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R237
            Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                237,
                @"[In tns:CopyItemSoapOut Message] [The CopyItemSoapOut WSDL message is defined as:] 
                    <wsdl:message name=""CopyItemSoapOut"">
                       <wsdl:part name=""CopyItemResult"" element=""tns:CopyItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R248");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R248
            // The schema is validated and the response of CopyItem operation is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                248,
                @"[In m:CopyItemResponse Element] [The CopyItemResponse element is defined as:]
                    <xs:element name=""CopyItemResponse""
                      type=""m:CopyItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R254");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R254
            Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                254,
                @"[In m:CopyItemResponseType Complex Type] [The CopyItemResponseType complex type is defined as:]
                    <xs:complexType name=""CopyItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1404");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1404
            Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                1404,
                @"[In tns:CopyItemSoapOut Message] The type of CopyItemResult is tns:CopyItemResponse (section 3.1.4.1.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R240");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R240
            // According to the schema, copyItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                240,
                @"[In tns:CopyItemSoapOut Message] [The part ""CopyItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1405");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1405
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1405,
                @"[In tns:CopyItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R241");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R241
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                241,
                @"[In tns:CopyItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for a response.");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(copyItemResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R220");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R220
            // The request of CopyItem operation is formed according to schema
            // And the response of CopyItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                220,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""CopyItem""] Copies items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R221");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R221
            // The request of CopyItem operation is formed according to schema
            // And the response of CopyItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                221,
                @"[In CopyItem Operation] The CopyItem operation copies items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R244");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R244
            // This requirement can be captured directly, since CopyItemResponse is the response of a CopyItem operation request.
            Site.CaptureRequirement(
                244,
                @"[In Elements] [The element ""CopyItemResponse""] Specifies a response to a CopyItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R247");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R247
            // This requirement can be captured directly, since CopyItemResponse is the response of a CopyItem operation request.
            Site.CaptureRequirement(
                247,
                @"[In m:CopyItemResponse Element] The CopyItemResponse element specifies a response to a CopyItem operation request.");

            foreach (ItemInfoResponseMessageType info in copyItemResponse.ResponseMessages.Items)
            {
                this.VerifyItemInfoResponseMessageType(info);
            }
        }
        #endregion

        #region Verify CreateItemResponseType Structure
        /// <summary>
        /// Verify the CreateItemResponseTypes structure.
        /// </summary>
        /// <param name="createItemResponse">A CreateItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyCreateItemResponse(CreateItemResponseType createItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R259");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R259
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                259,
                @"[In CreateItem Operation] The following is the WSDL port type specification for the CreateItem operation: 
                    <wsdl:operation name=""CreateItem"" xmlns:wsdl=""http://schemas.xmlsoap.org/wsdl/"">
                         <wsdl:input message=""tns:CreateItemSoapIn"" />
                         <wsdl:output message=""tns:CreateItemSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R260");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R260
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                260,
                @"[In CreateItem Operation] The following is the WSDL binding specification for the CreateItem operation: 
                    <wsdl:operation name=""CreateItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/CreateItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:CreateItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:CreateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:CreateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:header message=""tns:CreateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""CreateItemResult"" use=""literal""/>
                          <soap:header message=""tns:CreateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R276");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R276
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                276,
                @"[In tns:CreateItemSoapOut Message] [The CreateItemSoapOut WSDL message  is defined as:] 
                    <wsdl:message name=""CreateItemSoapOut"">
                       <wsdl:part name=""CreateItemResult"" element=""tns:CreateItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R287");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R287
            // The schema is validated and the response of CreateItem operation is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                287,
                @"[In m:CreateItemResponse Element] [The CreateItemResponse element is defined as:]
                    <xs:element name=""CreateItemResponse""
                      type=""m:CreateItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R293");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R293
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                293,
                @"[In m:CreateItemResponseType Complex Type] [The CreateItemResponseType complex type is defined as:] 
                    <xs:complexType name=""CreateItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1411");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1411
            Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                1411,
                @"[In tns:CreateItemSoapOut Message] The type of CreateItemResult is tns:CreateItemResponse (section 3.1.4.2.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R279");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R279
            // According to the schema, createItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                279,
                @"[In tns:CreateItemSoapOut Message] [The part ""CreateItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1412");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1412
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1412,
                @"[In tns:CreateItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R280");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R280
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                280,
                @"[In tns:CreateItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for the response to a CreateItem operation request.");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(createItemResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R216");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R216
            // The request of CreateItem operation is formed according to schema
            // And the response of CreateItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                216,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""CreateItem""] Creates items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R258");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R258
            // The request of CreateItem operation is formed according to schema
            // And the response of CreateItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                258,
                @"[In CreateItem Operation] The CreateItem operation creates items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R283");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R283
            // This requirement can be captured directly, since CreateItemResponse is the response of a CreateItem operation request.
            Site.CaptureRequirement(
                283,
                @"[In Elements] [The element ""CreateItemResponse""] Specifies a response to a CreateItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R286");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R286
            // This requirement can be captured directly, since CreateItemResponse is the response of a CreateItem operation request.
            Site.CaptureRequirement(
                286,
                @"[In m:CreateItemResponse Element] The CreateItemResponse element specifies a response to a CreateItem operation request.");

            foreach (ItemInfoResponseMessageType info in createItemResponse.ResponseMessages.Items)
            {
                this.VerifyItemInfoResponseMessageType(info);
            }
        }
        #endregion

        #region Verify DeleteItemResponseType Structure
        /// <summary>
        /// Verify the DeleteItemResponseType structure.
        /// </summary>
        /// <param name="deleteItemResponse">A DeleteItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyDeleteItemResoponse(DeleteItemResponseType deleteItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R304");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R304
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                304,
                @"[In DeleteItem Operation] The following is the WSDL port type specification for the DeleteItem operation: 
                    <wsdl:operation name=""DeleteItem"">
                       <wsdl:input message=""tns:DeleteItemSoapIn""/>
                       <wsdl:output message=""tns:DeleteItemSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R305");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R305
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                305,
                @"[In DeleteItem Operation] The following is the WSDL binding specification for the DeleteItem operation: 
                    <wsdl:operation name=""DeleteItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/DeleteItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:DeleteItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:DeleteItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:DeleteItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""DeleteItemResult"" use=""literal""/>
                          <soap:header message=""tns:DeleteItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R319");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R319
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                319,
                @"[In tns:DeleteItemSoapOut Message] [The DeleteItemSoapOut WSDL message is defined as:]
                    <wsdl:message name=""DeleteItemSoapOut"">
                       <wsdl:part name=""DeleteItemResult"" element=""tns:DeleteItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R330");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R330
            // The schema is validated and the response of DeleteItem operation is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                330,
                @"[In m:DeleteItemResponse Element] [The DeleteItemResponse element is defined as:]
                    <xs:element name=""DeleteItemResponse""
                     type=""m:DeleteItemResponseType""
                    />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R336");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R336
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                336,
                @"[In m:DeleteItemResponseType Complex Type] [The DeleteItemResponseType complex type is defined as:]
                    <xs:complexType name=""DeleteItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1421");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1421
            Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                1421,
                @"[In tns:DeleteItemSoapOut Message] The type of DeleteItemResult is tns:DeleteItemResponse (section 3.1.4.3.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R322");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R322
            // According to the schema, deleteItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                322,
                @"[In tns:DeleteItemSoapOut Message] [The part ""DeleteItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1422");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1422
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1422,
                @"[In tns:DeleteItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R323");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R323
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                323,
                @"[In tns:DeleteItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R217");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R217
            // The request of DeleteItem operation is formed according to schema
            // And the response of DeleteItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                217,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""DeleteItem""] Deletes items on the server.");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(deleteItemResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R303");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R303
            // The request of DeleteItem operation is formed according to schema
            // And the response of DeleteItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                303,
                @"[In DeleteItem Operation] The DeleteItem operation deletes items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R326");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R326
            // This requirement can be captured directly, since DeleteItemResponse is the response of a DeleteItem operation request.
            Site.CaptureRequirement(
                326,
                @"[In Elements] [The element ""DeleteItemResponse""] Specifies a response to a single DeleteItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R329");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R329
            // This requirement can be captured directly, since DeleteItemResponse is the response of a DeleteItem operation request.
            Site.CaptureRequirement(
                329,
                @"[In m:DeleteItemResponse Element] The DeleteItemResponse element specifies a response to a single DeleteItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R76");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R76
            // The DeleteType element as DisposalType is a required element, this requirement can be captured directly if the schema is validated.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                76,
                @"[In t:DisposalType Simple Type] The type [DisposalType] is defined as follow:
                    <xs:simpleType name=""DisposalType"">
                        <xs:restriction base=""xs:string"">
                            <xs:enumeration value=""HardDelete""/>
                            <xs:enumeration value=""MoveToDeletedItems""/>
                            <xs:enumeration value=""SoftDelete""/>
                        </xs:restriction>
                    </xs:simpleType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R343");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R343
            // This requirement can be captured directly after MS-OXWSCDATA_R76
            Site.CaptureRequirement(
                343,
                @"[In m:DeleteItemType Complex Type] [The attribute ""DeleteType""] Specifies an enumeration value that describes how an item is to be deleted.");
        }
        #endregion

        #region Verify GetItemResponseType Structure
        /// <summary>
        /// Verify the GetItemResponseType structure.
        /// </summary>
        /// <param name="getItemResponse">A GetItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyGetItemResponse(GetItemResponseType getItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R347");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R347
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                347,
                @"[In GetItem Operation] The following is the WSDL port type specification for the GetItem operation: 
                    <wsdl:operation name=""GetItem"">
                       <wsdl:input message=""tns:GetItemSoapIn""/>
                       <wsdl:output message=""tns:GetItemSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R348");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R348
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                348,
                @"[In GetItem Operation] The following is the WSDL binding specification for the GetItem operation: 
                    <wsdl:operation name=""GetItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/GetItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:GetItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:GetItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:GetItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:header message=""tns:GetItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                          <soap:header message=""tns:GetItemSoapIn"" part=""DateTimePrecision"" use=""literal"" />
                    <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""GetItemResult"" use=""literal""/>
                          <soap:header message=""tns:GetItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R364");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R364
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                364,
                @"[In tns:GetItemSoapOut Message] [The GetItemSoapOut WSDL message is defined as:] 
                    <wsdl:message name=""GetItemSoapOut"">
                       <wsdl:part name=""GetItemResult"" element=""tns:GetItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R375");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R375
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                375,
                @"[In m:GetItemResponse Element] [The GetItemResponse element is defined as:]
                    <xs:element name=""GetItemResponse""
                      type=""m:GetItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R381");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R381
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                381,
                @"[In m:GetItemResponseType Complex Type] [The GetItemResponseType complex type is defined as:] 
                    <xs:complexType name=""GetItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1434");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1434
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                1434,
                @"[In tns:GetItemSoapOut Message] The type of GetItemResult is tns:GetItemResponse (section 3.1.4.4.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R386");

            // The ItemShape element is a required element in request, the response shape is determined by this element, and if the response is not null,
            // this requirement is validated.
            Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                386,
                @"[In m:GetItemType Complex Type] [The element ""ItemShape""] Specifies the response shape of a GetItem operation response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R367");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R367
            // According to the schema, getItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                367,
                @"[In tns:GetItemSoapOut Message] [The part ""GetItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1435");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1435
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1435,
                @"[In tns:GetItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R368");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R368
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                368,
                @"[In tns:GetItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for a response.");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(getItemResponse);

            foreach (ItemInfoResponseMessageType info in getItemResponse.ResponseMessages.Items)
            {
                this.VerifyItemInfoResponseMessageType(info);

                if (Common.IsRequirementEnabled(2313, this.Site)
                    && this.exchangeServiceBinding.TimeZoneContext != null
                    && this.exchangeServiceBinding.TimeZoneContext.TimeZoneDefinition != null
                    && this.exchangeServiceBinding.TimeZoneContext.TimeZoneDefinition.Id == "Pacific Standard Time")
                {
                    string innerXml = this.exchangeServiceBinding.LastRawResponseXml.CreateNavigator().InnerXml;
                    string temp = innerXml.Substring(innerXml.IndexOf("DateTimeCreated"));
                    string dateTimeCreated = temp.Substring(temp.IndexOf(">") + 1, temp.IndexOf("<") - temp.IndexOf(">") - 1);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R361");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R361
                    Site.CaptureRequirementIfIsTrue(
                        dateTimeCreated.Contains(System.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(info.Items.Items[0].DateTimeCreated, "Pacific Standard Time").GetDateTimeFormats('s')[0]),
                        361,
                        @"[In tns:GetItemSoapIn Message] [The part ""TimeZoneContext""] Specifies a SOAP header that identifies the time zone to be used for all responses from the server.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2313");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2313
                    Site.CaptureRequirementIfIsTrue(
                        dateTimeCreated.Contains(System.TimeZoneInfo.ConvertTimeBySystemTimeZoneId(info.Items.Items[0].DateTimeCreated, "Pacific Standard Time").GetDateTimeFormats('s')[0]),
                        2313,
                        @"[In Appendix C: Product Behavior] Implementation does convert the times in GetItem response even if the TimeZoneContext SOAP header is set in request. (Exchange 2010 and above follow this behavior.) ");
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R215");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R215
            // The request of GetItem operation is formed according to schema
            // And the response of GetItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                215,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""GetItem""] Gets items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R346");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R346
            // Items are returned in ItemInfoResponseMessageType, if ItemInfoResponseMessageType is verified, this requirement can be verified.
            Site.CaptureRequirement(
                346,
                @"[In GetItem Operation] The GetItem operation gets items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R371");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R371
            // This requirement can be captured directly, since DeleteItemResponse is the response of a DeleteItem operation request.
            Site.CaptureRequirement(
                371,
                @"[In Elements] [The element ""GetItemResponse""] Specifies a response to a GetItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R374");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R374
            // This requirement can be captured directly, since DeleteItemResponse is the response of a DeleteItem operation request.
            Site.CaptureRequirement(
                374,
                @"[In m:GetItemResponse Element] The GetItemResponse element specifies a response to a GetItem operation request.");
        }
        #endregion

        #region Verify MarkAllItemsAsReadResponseTypes Structure
        /// <summary>
        /// Verify the MarkAllItemsAsReadResponseTypes structure.
        /// </summary>
        /// <param name="markAllItemsAsReadResponseType">A MarkAllItemsAsReadResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyMarkAllItemsAsReadResponse(MarkAllItemsAsReadResponseType markAllItemsAsReadResponseType, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1168");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1168
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1168,
                @"[In MarkAllItemsAsRead Operation] The following is the WSDL port type specification for the MarkAllItemsAsRead operation: <wsdl:operation name=""MarkAllItemsAsRead"">
                      <wsdl:input message=""tns:MarkAllItemsAsReadSoapIn""/>
                      <wsdl:output message=""tns:MarkAllItemsAsReadSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1170");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1170
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1170,
                @"[In MarkAllItemsAsRead Operation] The following is the WSDL binding specification for the MarkAllItemsAsRead operation: <wsdl:operation name=""MarkAllItemsAsRead"">
                      <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MarkAllItemsAsRead""/>
                      <wsdl:input>
                        <soap:header message=""tns:MarkAllItemsAsReadSoapIn"" part=""Impersonation"" use=""literal""/>
                        <soap:header message=""tns:MarkAllItemsAsReadSoapIn"" part=""MailboxCulture"" use=""literal""/>
                        <soap:header message=""tns:MarkAllItemsAsReadSoapIn"" part=""RequestVersion"" use=""literal""/>
                        <soap:body parts=""request"" use=""literal""/>
                      </wsdl:input>
                      <wsdl:output>
                        <soap:body parts=""MarkAllItemsAsReadResult"" use=""literal""/>
                        <soap:header message=""tns:MarkAllItemsAsReadSoapOut"" part=""ServerVersion"" use=""literal""/>
                      </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1187");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1187
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1187,
                @"[In tns:MarkAllItemsAsReadSoapOut Message] The message [MarkAllItemsAsReadSoapOut] is defined as follow: <wsdl:message name=""MarkAllItemsAsReadSoapOut"">
                      <wsdl:part name=""MarkAllItemsAsReadResult"" element=""tns:MarkAllItemsAsReadResponse""/>
                      <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1201");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1201
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1201,
                @"[In m:MarkAllItemsAsReadResponse Element] The element [MarkAllItemsAsReadResponse] is defined as follow: <xs:element name=""MarkAllItemsAsReadResponse"" type=""m:MarkAllItemsAsReadResponseType""/>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1213");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1213
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1213,
                @"[In m:MarkAllItemsAsReadResponseType Complex Type] The complex type [MarkAllItemsAsReadResponseType] is defined as follow: <xs:complexType name=""MarkAllItemsAsReadResponseType"">
                  <xs:complexContent>
                    <xs:extension base=""m:BaseResponseMessageType""/>
                  </xs:complexContent>
                </xs:complexType>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(markAllItemsAsReadResponseType);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1442");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1442
            Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponseType,
                1442,
                @"[In tns:MarkAllItemsAsReadSoapOut Message] The type of MarkAllItemsAsReadResult is tns:MarkAllItemsAsReadResponse (section 3.1.4.5.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1190");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1190
            // According to the schema, markAllItemsAsReadResponseType is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                1190,
                @"[In tns:MarkAllItemsAsReadSoapOut Message] [The part name ""MarkAllItemsAsReadResult""] 
                Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1443");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1443
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1443,
                @"[In tns:MarkAllItemsAsReadSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1191");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1191
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                1191,
                @"[In tns:MarkAllItemsAsReadSoapOut Message] [The part name ""ServerVersion""] 
                    Specifies a SOAP header that identifies the server version for the response to a MarkAllItemsAsRead operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1195");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1195
            // This requirement can be captured directly, since MarkAllItemsAsReadResponse is the response of a MarkAllItemsAsRead operation request.
            Site.CaptureRequirement(
                1195,
                @"[In Elements] [The element ""MarkAllItemsAsReadResponse""] Specifies a response to a MarkAllItemsAsRead operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1200");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1200
            // The request of MarkAllItemsAsRead operation is formed according to schema
            // And the response of MarkAllItemsAsRead operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                1200,
                @"[In m:MarkAllItemsAsReadResponse Element] The MarkAllItemsAsReadResponse element specifies a response to a MarkAllItemsAsRead operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1157");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1157
            // The request of MarkAllItemsAsRead operation is formed according to schema
            // And the response of MarkAllItemsAsRead operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                1157,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""MarkAllItemsAsRead""] Marks all the items in a folder as read.");
        }
        #endregion

        #region Verify MarkAsJunkResponseType Structure
        /// <summary>
        /// Verify the MarkAsJunkResponseType structure.
        /// </summary>
        /// <param name="markAsJunkReponse">A MarkAsJunkResponseType instance</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyMarkAsJunkResponse(MarkAsJunkResponseType markAsJunkReponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1791");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1791
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1791,
                @"[In MarkAsJunk Operation] The following is the WSDL port type specification of the MarkAsJunk operation.
                    <wsdl:operation name=""MarkAsJunk"">
                      <wsdl:input message=""tns:MarkAsJunkSoapIn""/>
                      <wsdl:output message=""tns:MarkAsJunkSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1792");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1792
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1792,
                @"[In MarkAsJunk Operation] The following is the WSDL binding specification of the MarkAsJunk operation
                    <wsdl:operation name=""MarkAsJunk"">
                      <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MarkAsJunk""/>
                      <wsdl:input>
                        <soap:header message=""tns:MarkAsJunkSoapIn"" part=""Impersonation"" use=""literal""/>
                        <soap:header message=""tns:MarkAsJunkSoapIn"" part=""MailboxCulture"" use=""literal""/>
                        <soap:header message=""tns:MarkAsJunkSoapIn"" part=""RequestVersion"" use=""literal""/>
                        <soap:body parts=""request"" use=""literal""/>
                      </wsdl:input>
                      <wsdl:output>
                        <soap:body parts=""MarkAsJunkResult"" use=""literal""/>
                        <soap:header message=""tns:MarkAsJunkSoapOut"" part=""ServerVersion"" use=""literal""/>
                      </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1809");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1809
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1809,
                @"[In tns:MarkAsJunkSoapOut Message] [The MarkAsJunkSoapOut WSDL message is defined as:] <wsdl:message name=""MarkAsJunkSoapOut"">
                      <wsdl:part name=""MarkAsJunkResult"" element=""tns:MarkAsJunkResponse""/>
                      <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1812");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1812
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1812,
                @"[In tns:MarkAsJunkSoapOut Message] The type of MarkAsJunkResult is tns:MarkAsJunkResponse (section 3.1.4.6.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1813");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1813
            // According to the schema, markAsJunkReponse is the SOAP body of a response message returned by server, this requirement can be verified directly.
            Site.CaptureRequirement(
                1813,
                @"[In tns:MarkAsJunkSoapOut Message] [The part ""MarkAsJunkResult""] Specifies the SOAP body of the response to a MarkAsJunk operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1814");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1814
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1814,
                @"[In tns:MarkAsJunkSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1815");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1815
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, this requirement can be verified directly.
            Site.CaptureRequirement(
                1815,
                @"[In tns:MarkAsJunkSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for the response to a MarkAsJunk operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1818");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1818
            // MarkAsJunkResponse is the response of MarkAsJunk operation, this requirement can be captured directly.
            Site.CaptureRequirement(
                1818,
                @"[In Elements] [The element ""MarkAsJunkResponse"" is] The result data for the MarkAsJunk WSDL operation.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1821");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1821
            // MarkAsJunkResponse is the response of MarkAsJunk operation, this requirement can be captured directly.
            Site.CaptureRequirement(
                1821,
                @"[In m:MarkAsJunkResponse Element] The MarkAsJunkResponse element specifies the response for a MarkAsJunk operation, as specified in section 3.1.4.6.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1822");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1822
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1822,
                @"[In m:MarkAsJunkResponse Element] [The MarkAsJunkResponse element is defined as:] <xs:element name=""MarkAsJunkResponse"" type=""m:MarkAsJunkResponseType"" />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1845");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1845
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1845,
                @"[In m:MarkAsJunkResponseType Complex Type] [The MarkAsJunkResponseType Complex Type is defined as:] <xs:complexType name=""MarkAsJunkResponseType"">
                    <xs:complexContent>
                      <xs:extension base=""m:BaseResponseMessageType"" />
                    </xs:complexContent>
                  </xs:complexType");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(markAsJunkReponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1848");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1848
            Site.CaptureRequirementIfIsNotNull(
                markAsJunkReponse,
                1848,
                @"[In m:MarkAsJunkResponseMessageType Complex Type] [The MarkAsJunkResponseMessageType Complex Type is defined as:] <xs:complexType name=""MarkAsJunkResponseMessageType"">
                    <xs:complexContent>
                      <xs:extension base=""m:ResponseMessageType"">
                        <xs:sequence>
                          <xs:sequence>
                            <xs:element name=""MovedItemId"" type=""t:ItemIdType"" minOccurs=""0"" maxOccurs=""1""/>
                          </xs:sequence>
                        </xs:sequence>
                      </xs:extension>
                    </xs:complexContent>
                  </xs:complexType>");
        }
        #endregion

        #region Verify MoveItemResponseType Structure
        /// <summary>
        /// Verify the MoveItemResponseType structure.
        /// </summary>
        /// <param name="moveItemResponse">A MoveItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyMoveItemResponse(MoveItemResponseType moveItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R389");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R389
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                389,
                @"[In MoveItem Operation] The following is the WSDL port type specification for the MoveItem operation: 
                    <wsdl:operation name=""MoveItem"">
                       <wsdl:input message=""tns:MoveItemSoapIn""/>
                       <wsdl:output message=""tns:MoveItemSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R390");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R390
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                390,
                @"[In MoveItem Operation] The following is the WSDL binding specification for the MoveItem operation: 
                    <wsdl:operation name=""MoveItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/MoveItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:MoveItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:MoveItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                           <soap:header message=""tns:MoveItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""MoveItemResult"" use=""literal""/>
                          <soap:header message=""tns:MoveItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2182");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2182
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                2182,
                @" [In tns:MoveItemSoapOut Message] [The MoveItemSoapOut WSDL message is defined as:]
                   <wsdl:message name=""MoveItemSoapOut"">
                       <wsdl:part name=""MoveItemResult"" element=""tns:MoveItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                   </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R414");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R414
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                414,
                @"[In m:MoveItemResponse Element] [The MoveItemResponse element is defined as:]
                    <xs:element name=""MoveItemResponse""
                      type=""m:MoveItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1448");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1448
            Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                1448,
                @"[In tns:MoveItemSoapOut Message] The type of MoveItemResult is tns:MoveItemResponse (section 3.1.4.7.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R406");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R406
            // According to the schema, moveItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                406,
                @"[In tns:MoveItemSoapOut Message] [The part ""MoveItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1449");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1449
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1449,
                @"[In tns:MoveItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R407");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R407
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                407,
                @"[In tns:MoveItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for the response to a MoveItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R420");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R420
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                420,
                @"[In m:MoveItemResponseType Complex Type] [The MoveItemResponseType complex type is defined as:]
                    <xs:complexType name=""MoveItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(moveItemResponse);

            foreach (ItemInfoResponseMessageType info in moveItemResponse.ResponseMessages.Items)
            {
                this.VerifyItemInfoResponseMessageType(info);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R219");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R219
            // The request of MoveItem operation is formed according to schema
            // And the response of MoveItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                219,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""MoveItem""] Moves items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R388");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R388
            // The request of MoveItem operation is formed according to schema
            // And the response of MoveItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                388,
                @"[In MoveItem Operation] The MoveItem operation moves items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R413");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R413
            // This requirement can be captured directly, since MoveItemResponse is the response of a MoveItem operation request.
            Site.CaptureRequirement(
                413,
                @"[In m:MoveItemResponse Element] The MoveItemResponse element specifies a response to a MoveItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R410");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R410
            // This requirement can be captured directly, since MoveItemResponse is the response of a MoveItem operation request.
            Site.CaptureRequirement(
                410,
                @"[In Elements] [The element ""MoveItemResponse""] Specifies a response to a MoveItem operation request.");
        }
        #endregion

        #region Verify SendItemResponseType Structure
        /// <summary>
        /// Verify the SendItemResponseType structure.
        /// </summary>
        /// <param name="sendItemResponse">A SendItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifySendItemResponse(SendItemResponseType sendItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R425");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R425
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                425,
                @"[In SendItem Operation] The following is the WSDL port type specification for the SendItem operation: 
                    <wsdl:operation name=""SendItem"">
                       <wsdl:input message=""tns:SendItemSoapIn"" />
                       <wsdl:output message=""tns:SendItemSoapOut"" />
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R426");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R426
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                426,
                @"[In SendItem Operation] The following is the WSDL binding specification for the SendItem operation:
                    <wsdl:operation name=""SendItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/SendItem"" />
                       <wsdl:input>
                          <soap:header message=""tns:SendItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:SendItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:SendItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal"" />
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""SendItemResult"" use=""literal"" />
                          <soap:header message=""tns:SendItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R440");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R440
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                440,
                @"[In tns:SendItemSoapOut Message] [The SendItemSoapOut WSDL message is defined as:]
                    <wsdl:message name=""SendItemSoapOut"">
                       <wsdl:part name=""SendItemResult"" element=""tns:SendItemResponse"" />
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1454");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1454
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                1454,
                @"[In tns:SendItemSoapOut Message] The type of SendItemResult is tns:SendItemResponse (section 3.1.4.8.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R443");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R443
            // According to the schema, sendItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                443,
                @"[In tns:SendItemSoapOut Message] [The part ""SendItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1455");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1455
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1455,
                @"[In tns:SendItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R444");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R444
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                444,
                @"[In tns:SendItemSoapOut Message] [The part ""ServerVersion""] Specifies a SOAP header that identifies the server version for the response to a SendItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R451");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R451
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                451,
                @"[In m:SendItemResponse Element] [The SendItemResponse element is defined as:]
                    <xs:element name=""SendItemResponse""
                      type=""m:SendItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R457");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R457
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                457,
                @"[In m:SendItemResponseType Complex Type] [The SendItemResponseType complex type is defined as:]
                    <xs:complexType name=""SendItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(sendItemResponse);

            foreach (ResponseMessageType message in sendItemResponse.ResponseMessages.Items)
            {
                this.VerifyResponseMessageType(message);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R450");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R450
            // This requirement can be captured directly, since SendItemResponse is the response of a SendItem operation request.
            Site.CaptureRequirement(
                450,
                @"[In m:SendItemResponse Element] The SendItemResponse element specifies a response to a SendItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R447");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R447
            // This requirement can be captured directly, since SendItemResponse is the response of a SendItem operation request.
            Site.CaptureRequirement(
                447,
                @"[In Elements] [The element ""SentItemResponse""] Specifies a response to a SendItem operation request.");

            this.VerifySendItemResponseTypeSchema(sendItemResponse);
        }
        #endregion

        #region Verify UpdateItemResponseType Structure
        /// <summary>
        /// Verify the UpdateItemResponseType structure.
        /// </summary>
        /// <param name="updateItemResponse">An UpdateItemResponseType instance.</param>
        /// <param name="isSchemaValidated">Indicate whether schema is verified.</param>
        private void VerifyUpdateItemResponse(UpdateItemResponseType updateItemResponse, bool isSchemaValidated)
        {
            Site.Assert.IsTrue(isSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R467");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R467
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                467,
                @"[In UpdateItem Operation] The following is the WSDL port type specification for the UpdateItem operation:
                    <wsdl:operation name=""UpdateItem"">
                       <wsdl:input message=""tns:UpdateItemSoapIn""/>
                       <wsdl:output message=""tns:UpdateItemSoapOut""/>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R468");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R468
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                468,
                @"[In UpdateItem Operation] The following is the WSDL binding specification for the UpdateItem operation: 
                    <wsdl:operation name=""UpdateItem"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/UpdateItem""/>
                       <wsdl:input>
                          <soap:header message=""tns:UpdateItemSoapIn"" part=""Impersonation"" use=""literal""/>
                          <soap:header message=""tns:UpdateItemSoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:UpdateItemSoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:header message=""tns:UpdateItemSoapIn"" part=""TimeZoneContext"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal""/>
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""UpdateItemResult"" use=""literal""/>
                          <soap:header message=""tns:UpdateItemSoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R484");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R484
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                484,
                @"[In tns:UpdateItemSoapOut Message] [The UpdateItemSoapOut WSDL message is defined as:]
                    <wsdl:message name=""UpdateItemSoapOut"">
                       <wsdl:part name=""UpdateItemResult"" element=""tns:UpdateItemResponse""/>
                       <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                    </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1464");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1464
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                1464,
                @"[In tns:UpdateItemSoapOut Message] The type of UpdateItemResult is tns:UpdateItemResponse (section 3.1.4.9.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R487");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R487
            // According to the schema, updateItemResponse is the SOAP body of a response message returned by server, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                487,
                @"[In tns:UpdateItemSoapOut Message] [The part ""UpdateItemResult""] Specifies the SOAP body of a response message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1465");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1465
            Site.CaptureRequirementIfIsNotNull(
                this.exchangeServiceBinding.ServerVersionInfoValue,
                1465,
                @"[In tns:UpdateItemSoapOut Message] The type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.5.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R488");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R488
            // According to the schema, ServerVersion is the SOAP header that contains the server version information, 
            // this requirement can be verified directly.
            Site.CaptureRequirement(
                488,
                @"[In tns:UpdateItemSoapOut Message] [The part ""ServerVersion"" with element ""t:ServerVersionInfo""] Specifies a SOAP header that identifies the server version for a response to an UpdateItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R495");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R495
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                495,
                @"[In m:UpdateItemResponse Element] [The UpdateItemResponse element is defined as:]
                    <xs:element name=""UpdateItemResponse""
                      type=""m:UpdateItemResponseType""
                     />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R509");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R509
            Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                509,
                @"[In m:UpdateItemResponseType Complex Type] [The UpdateItemResponseType complex type is defined as:] 
                    <xs:complexType name=""UpdateItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(updateItemResponse);

            foreach (UpdateItemResponseMessageType message in updateItemResponse.ResponseMessages.Items)
            {
                this.VerifyUpdateItemResponseMessageType(message);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R218");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R218
            // The request of UpdateItem operation is formed according to schema
            // And the response of UpdateItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                218,
                @"[In Message Processing Events and Sequencing Rules] [The operation ""UpdateItem""] Updates items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R466");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R466
            // The request of UpdateItem operation is formed according to schema
            // And the response of UpdateItem operation is returned by server according to schema
            // This requirement can be verified directly.
            Site.CaptureRequirement(
                466,
                @"[In UpdateItem Operation] The UpdateItem operation updates items on the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R491");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R491
            // This requirement can be captured directly, since UpdateItemResponse is the response of a UpdateItem operation request.
            Site.CaptureRequirement(
                491,
                @"[In Elements] [The element ""UpdateItemResponse""] Specifies a response to an UpdateItem operation request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R494");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R494
            // This requirement can be captured directly, since UpdateItemResponse is the response of a UpdateItem operation request.
            Site.CaptureRequirement(
                494,
                @"[In m:UpdateItemResponse Element] The UpdateItemResponse element specifies a response to an UpdateItem operation request.");
        }
        #endregion
        #endregion

        #region Verify Section 2: Messages
        #region Verify Schema of SendItemResponseType
        /// <summary>
        /// Verify the SendItemResponseType structure in section 2.
        /// </summary>
        /// <param name="sendItemResponse">A SendItemResponseType instance.</param>
        private void VerifySendItemResponseTypeSchema(SendItemResponseType sendItemResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R50");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R50
            // The schema is validated and the response of SendItem operation is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                sendItemResponse,
                50,
                @"[In m:SendItemResponseType Complex Type] [The type SendItemResponseType is defined as follows:]
                     <xs:complexType name=""SendItemResponseType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:BaseResponseMessageType""
                         />
                      </xs:complexContent>
                    </xs:complexType>");

            // Verify the BaseResponseMessageType schema.
            this.VerifyBaseResponseMessageType(sendItemResponse);
        }
        #endregion

        #region Verify UpdateItemResponseMessageType Structure
        /// <summary>
        /// Verify the UpdateItemResponseMessageType structure.
        /// </summary>
        /// <param name="message">An UpdateItemResponseMessageType response message instance.</param>
        private void VerifyUpdateItemResponseMessageType(UpdateItemResponseMessageType message)
        {
            // UpdateItemResponseMessageType extends from the ItemInfoResponseMessageType.
            this.VerifyItemInfoResponseMessageType(message);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R59");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R59
            Site.CaptureRequirementIfIsNotNull(
                message,
                59,
                @"[In m:UpdateItemResponseMessageType Complex Type] [The type UpdateItemResponseMessageType is defined as follow:]
                    <xs:complexType name=""UpdateItemResponseMessageType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:ItemInfoResponseMessageType""
                        >
                          <xs:sequence>
                            <xs:element name=""ConflictResults""
                              type=""t:ConflictResultsType""
                              minOccurs=""0""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R63");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R63
            // If the schema is validated, this requirement can be captured.
            Site.CaptureRequirement(
                63,
                @"[In t:ConflictResultsType Complex Type] [The type ConflictResultsType is defined as follow:]
                    <xs:complexType name=""ConflictResultsType"">
                      <xs:sequence>
                        <xs:element name=""Count""
                          type=""xs:int""
                         />
                      </xs:sequence>
                    </xs:complexType>");
        }
        #endregion

        #region Verify ItemType Structure
        /// <summary>
        /// Verify the ItemType structure.
        /// </summary>
        /// <param name="item">An ItemType instance.</param>
        private void VerifyItemType(ItemType item)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R67");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R67
            // The schema is validated and the item is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                item,
                67,
                @"[In t:ItemType Complex Type] [The type ItemType is defined as follow:]
                    <xs:complexType name=""ItemType"">
                        <xs:sequence>
                            <xs:element name=""MimeContent"" type=""t:MimeContentType"" minOccurs=""0""/>
                            <xs:element name=""ItemId"" type=""t:ItemIdType"" minOccurs=""0""/>
                            <xs:element name=""ParentFolderId"" type=""t:FolderIdType"" minOccurs=""0""/>
                            <xs:element name=""ItemClass"" type=""t:ItemClassType"" minOccurs=""0""/>
                            <xs:element name=""Subject"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""Sensitivity"" type=""t:SensitivityChoicesType"" minOccurs=""0""/>
                            <xs:element name=""Body"" type=""t:BodyType"" minOccurs=""0""/>
                            <xs:element name=""Attachments"" type=""t:NonEmptyArrayOfAttachmentsType"" minOccurs=""0""/>
                            <xs:element name=""DateTimeReceived"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""Size"" type=""xs:int"" minOccurs=""0""/>
                            <xs:element name=""Categories"" type=""t:ArrayOfStringsType"" minOccurs=""0""/>
                            <xs:element name=""Importance"" type=""t:ImportanceChoicesType"" minOccurs=""0""/>
                            <xs:element name=""InReplyTo"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""IsSubmitted"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""IsDraft"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""IsFromMe"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""IsResend"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""IsUnmodified"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""InternetMessageHeaders"" type=""t:NonEmptyArrayOfInternetHeadersType"" minOccurs=""0""/>
                            <xs:element name=""DateTimeSent"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""DateTimeCreated"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""ResponseObjects"" type=""t:NonEmptyArrayOfResponseObjectsType"" minOccurs=""0""/>
                            <xs:element name=""ReminderDueBy"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""ReminderIsSet"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""ReminderNextTime"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""ReminderMinutesBeforeStart"" type=""t:ReminderMinutesBeforeStartType"" minOccurs=""0""/>
                            <xs:element name=""DisplayCc"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""DisplayTo"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""HasAttachments"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""ExtendedProperty"" type=""t:ExtendedPropertyType"" minOccurs=""0"" maxOccurs=""unbounded""/>
                            <xs:element name=""Culture"" type=""xs:language"" minOccurs=""0""/>
                            <xs:element name=""EffectiveRights"" type=""t:EffectiveRightsType"" minOccurs=""0""/>
                            <xs:element name=""LastModifiedName"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""LastModifiedTime"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""IsAssociated"" type=""xs:boolean"" minOccurs=""0""/>
                            <xs:element name=""WebClientReadFormQueryString"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""WebClientEditFormQueryString"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""ConversationId"" type=""t:ItemIdType"" minOccurs=""0""/>
                            <xs:element name=""UniqueBody"" type=""t:BodyType"" minOccurs=""0""/>
                            <xs:element name=""Flag"" type=""t:FlagType"" minOccurs=""0""/>
                            <xs:element name=""StoreEntryId"" type=""xs:base64Binary"" minOccurs=""0""/>
                            <xs:element name=""InstanceKey"" type=""xs:base64Binary"" minOccurs=""0""/>
                            <xs:element name=""NormalizedBody"" type=""t:BodyType"" minOccurs=""0""/>
                            <xs:element name=""EntityExtractionResult"" type=""t:EntityExtractionResultType"" minOccurs=""0"" />
                            <xs:element name=""PolicyTag"" type=""t:RetentionTagType"" minOccurs=""0""/>
                            <xs:element name=""ArchiveTag"" type=""t:RetentionTagType"" minOccurs=""0""/>
                            <xs:element name=""RetentionDate"" type=""xs:dateTime"" minOccurs=""0""/>
                            <xs:element name=""Preview"" type=""xs:string"" minOccurs=""0""/>
                            <xs:element name=""RightsManagementLicenseData"" type=""t:RightsManagementLicenseDataType"" minOccurs=""0"" />
                            <xs:element name=""NextPredictedAction"" type=""t:PredictedMessageActionType"" minOccurs=""0"" />
                            <xs:element name=""GroupingAction"" type=""t:PredictedMessageActionType"" minOccurs=""0""/>
                            <xs:element name=""BlockStatus"" type=""xs:boolean"" minOccurs=""0"" />
                            <xs:element name=""HasBlockedImages"" type=""xs:boolean"" minOccurs=""0"" />
                            <xs:element name=""TextBody"" type=""t:BodyType"" minOccurs=""0""/>
                            <xs:element name=""IconIndex"" type=""t:IconIndexType"" minOccurs=""0""/> 
                        </xs:sequence>
                    </xs:complexType>");

            if (item.Body != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1096");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1096
                // The schema is validated and the Body is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1096,
                    @"[In t:BodyType Complex Type] The type [BodyType] is defined as follow:
                        <xs:complexType name=""BodyType"">
                          <xs:simpleContent>
                            <xs:extension
                              base=""xs:string""
                            >
                              <xs:attribute name=""BodyType"" type=""t:BodyTypeType""/>
                              xs:attribute name=""IsTruncated"" type=""xs:boolean"" use=""optional""/>
                            </xs:extension>
                          </xs:simpleContent>
                        </xs:complexType>");
            }

            if (item.ExtendedProperty != null)
            {
                this.VerifyExtendedPropertyType(item.ExtendedProperty);
            }

            if (item.MimeContent != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R112");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R112
                // If schema is validated, and the MimeContent element is not null,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    112,
                    @"[In t:MimeContentType Complex Type] [The type MimeContentType is defined as follow:]
                        <xs:complexType name=""MimeContentType"">
                          <xs:simpleContent>
                            <xs:extension
                              base=""xs:string""
                            >
                              <xs:attribute name=""CharacterSet""
                                type=""xs:string""
                                use=""optional""
                               />
                            </xs:extension>
                          </xs:simpleContent>
                        </xs:complexType>");
            }

            if (item.Flag != null)
            {
                if (Common.IsRequirementEnabled(1271, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1271");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1271
                    // The schema is validated and the Flag is not null, so this requirement can be captured.
                    Site.CaptureRequirement(
                        1271,
                        @"[In Appendix C: Product Behavior] Implementation does support FlagType complex type which specifies a flag indicating status, start date, due date or completion date for an item. (Exchange and above follow this behavior.)
                            <xs:complexType name=""FlagType"">
                                <xs:sequence>
                                    <xs:element name=""FlagStatus"" type=""t:FlagStatusType"" minOccurs=""1"" maxOccurs=""1""/>
                                    <xs:element name=""StartDate"" type=""xs:dateTime"" minOccurs=""0""/>
                                    <xs:element name=""DueDate"" type=""xs:dateTime"" minOccurs=""0""/>
                                    <xs:element name=""CompleteDate"" type=""xs:dateTime"" minOccurs=""0""/>
                                </xs:sequence>
                            </xs:complexType>");
                }
            }

            if (item.EntityExtractionResult != null)
            {
                this.VerifyEntityExtractionResultType(item.EntityExtractionResult);
            }

            if (item.ResponseObjects != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R126");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R126
                // The schema is validated and the ResponseObjects is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    126,
                    @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] [The type NonEmptyArrayOfResponseObjectsType is defined as follow:]
                        <xs:complexType name=""NonEmptyArrayOfResponseObjectsType"">
                          <xs:choice
                            maxOccurs=""unbounded""
                            minOccurs=""0""
                          >
                            <xs:element name=""AcceptItem""
                              type=""t:AcceptItemType""
                             />
                            <xs:element name=""TentativelyAcceptItem""
                              type=""t:TentativelyAcceptItemType""
                             />
                            <xs:element name=""DeclineItem""
                              type=""t:DeclineItemType""
                             />
                            <xs:element name=""ReplyToItem""
                              type=""t:ReplyToItemType""
                             />
                            <xs:element name=""ForwardItem""
                              type=""t:ForwardItemType""
                             />
                            <xs:element name=""ReplyAllToItem""
                              type=""t:ReplyAllToItemType""
                             />
                            <xs:element name=""CancelCalendarItem""
                              type=""t:CancelCalendarItemType""
                             />
                            <xs:element name=""RemoveItem""
                              type=""t:RemoveItemType""
                             />
                            <xs:element name=""SuppressReadReceipt""
                              type=""t:SuppressReadReceiptType""
                             />
                            <xs:element name=""PostReplyItem""
                              type=""t:PostReplyItemType""
                             />
                            <xs:element name=""AcceptSharingInvitation""
                              type=""t:AcceptSharingInvitationType""
                             />
                            <xs:element name=""AddItemToMyCalendar""
                              type=""t:AddItemToMyCalendarType""
                             />
                            <xs:element name=""ProposeNewTime""
                              type=""t:ProposeNewTimeType""
                             />
                          </xs:choice>
                        </xs:complexType>");

                foreach (ResponseObjectType responseObject in item.ResponseObjects)
                {
                    this.VerifyResponseObjectType(responseObject);
                }
            }

            if (item.ItemId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R154");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R154
                // The schema is validated and the ItemId is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    154,
                    @"[In t:ItemIdType Complex Type] [The type ItemIdType is defined as follow:]
                        <xs:complexType name=""ItemIdType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:BaseItemIdType""
                            >
                              <xs:attribute name=""Id""
                                type=""xs:string""
                                use=""required""
                               />
                              <xs:attribute name=""ChangeKey""
                                type=""xs:string""
                                use=""optional""
                               />
                            </xs:extension>
                          </xs:complexContent>
                        </xs:complexType>");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1424");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1424
                // The schema is validated and the ItemId is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1424,
                @"[In t:BaseItemIdType Complex Type] 
                    The type [BaseItemIdType] is defined as follow:
                    <xs:complexType name=""BaseItemIdType""
                      abstract=""true""
                     />");
            }

            if (item.ImportanceSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R196");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R196
                // The schema is validated and the ImportanceSpecified is true, so this requirement can be captured.
                Site.CaptureRequirement(
                    196,
                    @"[In t:ImportanceChoicesType Simple Type] [The type ImportanceChoicesType is defined as follow:]
                        <xs:simpleType name=""ImportanceChoicesType"">
                          <xs:restriction
                            base=""xs:string""
                          >
                            <xs:enumeration
                              value=""High""
                             />
                            <xs:enumeration
                              value=""Low""
                             />
                            <xs:enumeration
                              value=""Normal""
                             />
                          </xs:restriction>
                        </xs:simpleType>");
            }

            if (item.ItemClass != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R202");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R202
                // The schema is validated and the ItemClass element is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    202,
                    @"[In t:ItemClassType Simple Type] [The type ItemClassType is defined as follow:]
                        <xs:simpleType name=""ItemClassType"">
                          <xs:restriction
                            base=""xs:string""
                           />
                        </xs:simpleType>");
            }

            if (item.ReminderMinutesBeforeStart != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R204");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R204
                // The schema is validated and the ReminderMinutesBeforeStart element is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    204,
                    @"[In t:ReminderMinutesBeforeStartType Simple Type] [The type ReminderMinutesBeforeStartType is defined as follow:]
                        <xs:simpleType name=""ReminderMinutesBeforeStartType"">
                          <xs:union>
                            <xs:simpleType
                              id=""ReminderMinutesBeforeStartType""
                            >
                              <xs:restriction
                                base=""xs:int""
                              >
                                <xs:minInclusive
                                  value=""0""
                                 />
                                <xs:maxInclusive
                                  value=""2629800""
                                 />
                              </xs:restriction>
                            </xs:simpleType>
                            <xs:simpleType
                              id=""ReminderMinutesBeforeStartMarkerType""
                            >
                              <xs:restriction
                                base=""xs:int""
                              >
                                <xs:minInclusive
                                  value=""1525252321""
                                 />
                                <xs:maxInclusive
                                  value=""1525252321""
                                 />
                              </xs:restriction>
                            </xs:simpleType>
                          </xs:union>
                        </xs:simpleType>");
            }

            if (item.Body != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R20");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R20
                // The schema is validated and the Body element is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    20,
                    @"[In t:BodyTypeType Simple Type] This type [BodyTypeType] is defined as follow:
                        <xs:simpleType name=""BodyTypeType"">
                          <xs:restriction
                            base=""xs:string""
                          >
                            <xs:enumeration
                              value=""HTML""
                             />
                            <xs:enumeration
                              value=""Text""
                             />
                          </xs:restriction>
                        </xs:simpleType>");
            }

            if (item.SensitivitySpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R783");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R783
                // The schema is validated and the SensitivitySpecified is true, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    783,
                    @"[In t:SensitivityChoicesType Simple Type] The type [SensitivityChoicesType] is defined as follow:
                        <xs:simpleType name=""SensitivityChoicesType"">
                            <xs:restriction base=""xs:string"">
                                <xs:enumeration value=""Confidential""/>
                                <xs:enumeration value=""Normal""/>
                                <xs:enumeration value=""Personal""/>
                                <xs:enumeration value=""Private""/>
                            </xs:restriction>
                        </xs:simpleType>");
            }

            if (item.Categories != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1081");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1081
                // The schema is validated and the Categories is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1081,
                    @"[In t:ArrayOfStringsType Complex Type] The type [ArrayOfStringsType] is defined as follow:
                        <xs:complexType name=""ArrayOfStringsType"">
                          <xs:sequence>
                            <xs:element name=""String""
                              type=""xs:string""
                              minOccurs=""0""
                              maxOccurs=""unbounded""
                             />
                          </xs:sequence>
                        </xs:complexType>");
            }

            if (item.EffectiveRights != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1129");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1129
                // The schema is validated and the EffectiveRights is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1129,
                    @"[In t:EffectiveRightsType Complex Type] The type [EffectiveRightsType] is defined as follow:
                         <xs:complexType name=""EffectiveRightsType"">
                          <xs:sequence>
                            <xs:element name=""CreateAssociated""
                              type=""xs:boolean""
                             />
                            <xs:element name=""CreateContents""
                              type=""xs:boolean""
                             />
                            <xs:element name=""CreateHierarchy""
                              type=""xs:boolean""
                             />
                            <xs:element name=""Delete""
                              type=""xs:boolean""
                             />
                            <xs:element name=""Modify""
                              type=""xs:boolean""
                             />
                            <xs:element name=""Read""
                              type=""xs:boolean""
                             />
                            <xs:element name=""ViewPrivateItems""
                              type=""xs:boolean""
                             />
                          </xs:sequence>
                        </xs:complexType>");
            }

            if (item.ParentFolderId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1165");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1165
                // The schema is validated and the ParentFolderId is not null, so this requirement can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1165,
                    @"[In t:FolderIdType Complex Type] The type [FolderIdType] is defined as follow:
                    <xs:complexType name=""FolderIdType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""t:BaseFolderIdType""
                        >
                          <xs:attribute name=""Id""
                            type=""xs:string""
                            use=""required""
                           />
                          <xs:attribute name=""ChangeKey""
                            type=""xs:string""
                            use=""optional""
                           />
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");
            }

            if (item.InternetMessageHeaders != null)
            {
                foreach (InternetHeaderType internetMessageHeader in item.InternetMessageHeaders)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1175");

                    // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1175
                    // The schema is validated and the internetMessageHeader is not null, so this requirement can be captured.
                    Site.CaptureRequirementIfIsNotNull(
                        internetMessageHeader,
                        "MS-OXWSCDATA",
                        1175,
                        @"[In t:InternetHeaderType Complex Type] The type [InternetHeaderType] is defined as follow:
                            <xs:complexType name=""InternetHeaderType"">
                              <xs:simpleContent>
                                <xs:extension
                                  base=""xs:string""
                                >
                                  <xs:attribute name=""HeaderName""
                                    type=""xs:string""
                                    use=""required""
                                   />
                                </xs:extension>
                              </xs:simpleContent>
                            </xs:complexType>");
                }
            }
        }
        #endregion
        #endregion

        #region Verify ResponseObjectType Structure
        /// <summary>
        /// Verify the ResponseObjectType structure
        /// </summary>
        /// <param name="responseObject">A ResponseObjectType instance.</param>
        private void VerifyResponseObjectType(ResponseObjectType responseObject)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1289");

            // If the responseObject element is not null and schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                responseObject,
                "MS-OXWSCDATA",
                1289,
                @"[In t:ResponseObjectType Complex Type] The type [ResponseObjectType] is defined as follow:
                    <xs:complexType name=""ResponseObjectType""
                      abstract=""true""
                    >
                      <xs:complexContent>
                        <xs:extension
                          base=""t:ResponseObjectCoreType""
                        >
                          <xs:attribute name=""ObjectName""
                            type=""xs:string""
                            use=""optional""
                           />
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            this.VerifyResponseObjectCoreType(responseObject);

            switch (responseObject.GetType().Name)
            {
                case "ForwardItemType":
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1173");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1173
                    // If the responseObject element is a ForwardItemType, it must not be null and schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSCDATA",
                        1173,
                        @"[In t:ForwardItemType Complex Type] The type [ForwardItemType] is defined as follow:
                        <xs:complexType name=""ForwardItemType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:SmartResponseType""
                             />
                          </xs:complexContent>
                        </xs:complexType>");
                    break;

                case "ReplyAllToItemType":
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1267");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1267
                    // If the responseObject element is a ReplyAllToItemType, it must not be null and schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSCDATA",
                        1267,
                        @"[In t:ReplyAllToItemType Complex Type] The type [ReplyAllToItemType] is defined as follow:
                        <xs:complexType name=""ReplyAllToItemType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:SmartResponseType""
                             />
                          </xs:complexContent>
                        </xs:complexType>");
                    break;
                case "ReplyToItemType":
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1276");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1276
                    // If the responseObject element is a ReplyToItemType, it must not be null and schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSCDATA",
                        1276,
                        @"[In t:ReplyToItemType Complex Type] The type [ReplyToItemType] is defined as follow:
                        <xs:complexType name=""ReplyToItemType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:SmartResponseType""
                             />
                          </xs:complexContent>
                        </xs:complexType>");
                    break;
                case "SuppressReadReceiptType":
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1295");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1295
                    // If the responseObject element is a SuppressReadReceiptType, it must not be null and schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSCDATA",
                        1295,
                        @"[In t:SuppressReadReceiptType Complex Type] The type [SuppressReadReceiptType] is defined as follow:
                        <xs:complexType name=""SuppressReadReceiptType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:ReferenceItemResponseType""
                             />
                          </xs:complexContent>
                        </xs:complexType>");
                    break;
            }
        }
        #endregion

        #region Verify ResponseObjectCoreType Structure
        /// <summary>
        /// Verify the ResponseObjectCoreType structure
        /// </summary>
        /// <param name="responseCoreObject">A ResponseObjectCoreType instance.</param>
        private void VerifyResponseObjectCoreType(ResponseObjectCoreType responseCoreObject)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1286");

            // If the responseCoreObject element is not null and schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                responseCoreObject,
                "MS-OXWSCDATA",
                1286,
                @"[In t:ResponseObjectCoreType Complex Type] The type [ResponseObjectCoreType] is defined as follow:
                    <xs:complexType name=""ResponseObjectCoreType""
                      abstract=""true""
                    >
                      <xs:complexContent>
                        <xs:extension
                          base=""t:MessageType""
                        >
                          <xs:sequence>
                            <xs:element name=""ReferenceItemId""
                              type=""t:ItemIdType""
                              minOccurs=""0""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");
        }
        #endregion

        #region Verify ExtendedPropertyType Structure
        /// <summary>
        /// Verify the ExtendedPropertyType structure
        /// </summary>
        /// <param name="extendedProperties">An array of ExtendedPropertyType instances.</param>
        private void VerifyExtendedPropertyType(ExtendedPropertyType[] extendedProperties)
        {
            foreach (ExtendedPropertyType property in extendedProperties)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R128");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R128
                Site.CaptureRequirementIfIsNotNull(
                    property,
                    "MS-OXWSXPROP",
                    128,
                    @"[In t:ExtendedPropertyType Complex Type] The ExtendedPropertyType is defined as following:
                        <xs:complexType name=""ExtendedPropertyType"">
                          <xs:sequence>
                            <xs:element name=""ExtendedFieldURI""
                              type=""t:PathToExtendedFieldType""
                             />
                            <xs:choice>
                              <xs:element name=""Value""
                                type=""xs:string""
                               />
                              <xs:element name=""Values""
                                type=""t:NonEmptyArrayOfPropertyValuesType""
                               />
                            </xs:choice>
                          </xs:sequence>
                        </xs:complexType>");

                if (property.Item is NonEmptyArrayOfPropertyValuesType)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R127");

                    // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R127
                    // The schema is validated and property.Item is not null, so this requirement can be captured.
                    Site.CaptureRequirement(
                        "MS-OXWSXPROP",
                        127,
                        @"[In t:NonEmptyArrayOfPropertyValuesType Complex Type] The NonEmptyArrayOfPropertyValuesType is defined as following:
                             <xs:complexType name=""NonEmptyArrayOfPropertyValuesType"">
                              <xs:choice>
                                <xs:element name=""Value""
                                  type=""xs:string""
                                  maxOccurs=""unbounded""
                                 />
                              </xs:choice>
                            </xs:complexType>");
                }

                if (property.ExtendedFieldURI != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R130");

                    // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R130
                    // The ExtendedFieldURI element is not null, and the schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSXPROP",
                        130,
                        @"[In t:PathToExtendedFieldType Complex Type] 
                        The PathToExtendedFieldType is defined as following:
                        <xs:complexType name=""PathToExtendedFieldType"">
                          <xs:complexContent>
                            <xs:extension
                              base=""t:BasePathToElementType""
                            >
                              <xs:attribute name=""DistinguishedPropertySetId""
                                type=""t:DistinguishedPropertySetType""
                                use=""optional""
                               />
                              <xs:attribute name=""PropertySetId""
                                type=""t:GuidType""
                                use=""optional""
                               />
                              <xs:attribute name=""PropertyTag""
                                type=""t:PropertyTagType""
                                use=""optional""
                               />
                              <xs:attribute name=""PropertyName""
                                type=""xs:string""
                                use=""optional""
                               />
                              <xs:attribute name=""PropertyId""
                                type=""xs:int""
                                use=""optional""
                               />
                              <xs:attribute name=""PropertyType""
                                type=""t:MapiPropertyTypeType""
                                use=""required""
                               />
                            </xs:extension>
                          </xs:complexContent>
                        </xs:complexType>");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R133");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R133
                    // The MapiPropertyTypeType element is a required element as defined in the schema, if the schema is validated,
                    // this requirement can be validated.
                    Site.CaptureRequirement(
                        "MS-OXWSXPROP",
                        133,
                        @"[In t:MapiPropertyTypeType Simple Type]The MapiPropertyTypeType is defined as following:
                        <xs:simpleType name=""MapiPropertyTypeType"">
                          <xs:restriction
                            base=""xs:string""
                          >
                            <xs:enumeration
                              value=""ApplicationTime""
                             />
                            <xs:enumeration
                              value=""ApplicationTimeArray""
                             />
                            <xs:enumeration
                              value=""Binary""
                             />
                            <xs:enumeration
                              value=""BinaryArray""
                             />
                            <xs:enumeration
                              value=""Boolean""
                             />
                            <xs:enumeration
                              value=""CLSID""
                             />
                            <xs:enumeration
                              value=""CLSIDArray""
                             />
                            <xs:enumeration
                              value=""Currency""
                             />
                            <xs:enumeration
                              value=""CurrencyArray""
                             />
                            <xs:enumeration
                              value=""Double""
                             />
                            <xs:enumeration
                              value=""DoubleArray""
                             />
                            <xs:enumeration
                              value=""Error""
                             />
                            <xs:enumeration
                              value=""Float""
                             />
                            <xs:enumeration
                              value=""FloatArray""
                             />
                            <xs:enumeration
                              value=""Integer""
                             />
                            <xs:enumeration
                              value=""IntegerArray""
                             />
                            <xs:enumeration
                              value=""Long""
                             />
                            <xs:enumeration
                              value=""LongArray""
                             />
                            <xs:enumeration
                              value=""Null""
                             />
                            <xs:enumeration
                              value=""Object""
                             />
                            <xs:enumeration
                              value=""ObjectArray""
                             />
                            <xs:enumeration
                              value=""Short""
                             />
                            <xs:enumeration
                              value=""ShortArray""
                             />
                            <xs:enumeration
                              value=""SystemTime""
                             />
                            <xs:enumeration
                              value=""SystemTimeArray""
                             />
                            <xs:enumeration
                              value=""String""
                             />
                            <xs:enumeration
                              value=""StringArray""
                             />
                          </xs:restriction>
                        </xs:simpleType>");

                    if (property.ExtendedFieldURI.PropertySetId != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R131");

                        // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R131
                        // The PropertySetId element is not null, the pattern is matched, and the schema is validated,
                        // this requirement can be validated.
                        Site.CaptureRequirementIfIsTrue(
                            Regex.IsMatch(property.ExtendedFieldURI.PropertySetId, @"[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}"),
                            "MS-OXWSXPROP",
                            131,
                            @"[In t:GuidType Simple Type] 
                                The GuidType is defined as following :
                                <xs:simpleType name=""GuidType"">
                                  <xs:restriction
                                    base=""xs:string""
                                  >
                                    <xs:pattern
                                      value=""[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}""
                                     />
                                  </xs:restriction>
                                </xs:simpleType>");
                    }

                    if (property.ExtendedFieldURI.PropertyTag != null)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R134");

                        // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R134
                        // The PropertyTag element is not null, the pattern is matched, and the schema is validated,
                        // this requirement can be validated.
                        Site.CaptureRequirementIfIsTrue(
                            Regex.IsMatch(property.ExtendedFieldURI.PropertyTag, @"(0x|0X)[0-9A-Fa-f]{1,4}"),
                            "MS-OXWSXPROP",
                            134,
                            @"[In t:PropertyTagType Simple Type] The PropertyTagType is defined as following:
                                <xs:simpleType name=""PropertyTagType"">
                                  <xs:union
                                    memberTypes=""xs:unsignedShort""
                                  >
                                    <xs:simpleType
                                      id=""HexPropertyTagType""
                                    >
                                      <xs:restriction
                                        base=""xs:string""
                                      >
                                        <xs:pattern
                                          value=""(0x|0X)[0-9A-Fa-f]{1,4}""
                                         />
                                      </xs:restriction>
                                    </xs:simpleType>
                                  </xs:union>
                                </xs:simpleType>");
                    }
                }
            }
        }
        #endregion

        #region Verify EntityExtractionResultType Structure
        /// <summary>
        /// Verify the EntityExtractionResultType structure
        /// </summary>
        /// <param name="entityExtractionResult">An array of EntityExtractionResultType instances.</param>
        private void VerifyEntityExtractionResultType(EntityExtractionResultType entityExtractionResult)
        {
            if (entityExtractionResult.Addresses != null)
            {
                if (Common.IsRequirementEnabled(1749, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1749");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1749
                    // The Addresses element is ArrayOfAddressEntitiesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1749,
                        @"[In Appendix C: Product Behavior] Implementation does use the ArrayOfAddressEntitiesType complex type which represents an array of address entities. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfAddressEntitiesType"">
                              <xs:sequence>
                               <xs:element name=""AddressEntity"" type=""t:AddressEntityType"" 
                                    minOccurs=""0"" maxOccurs=""unbounded""/>
                              </xs:sequence>
                            </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.Addresses.Length,
                    string.Format(
                    "There should be one address information in the entity extraction result, actual {0}.",
                    entityExtractionResult.Addresses.Length));

                if (Common.IsRequirementEnabled(1753, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1753");

                    // The address element is AddressEntityType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        entityExtractionResult.Addresses[0],
                        1753,
                        @"[In Appendix C: Product Behavior] Implementation does use the AddressEntityType complex type which extends the EntityType complex type, as specified by section 2.2.4.38. (Exchange 2013 and above follow this behavior.)
                        <xs:complexType name=""AddressEntityType"">
                          <xs:complexContent>
                            <xs:extension base=""t:EntityType"">
                              <xs:sequence>
                                <xs:element name=""Address"" type=""xs:string"" minOccurs=""0""/>
                              </xs:sequence>
                            </xs:extension>
                          </xs:complexContent>
                        </xs:complexType>. ");

                    this.VerifyEntityType(entityExtractionResult.Addresses[0]);
                }
            }

            if (entityExtractionResult.MeetingSuggestions != null)
            {
                if (Common.IsRequirementEnabled(1275, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1275");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1275
                    // The MeetingSuggestions element is ArrayOfMeetingSuggestionsType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1275,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfMeetingSuggestionsType complex type which specifies an array of meeting suggestions. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""ArrayOfMeetingSuggestionsType"">
                                <xs:sequence>
                                    <xs:element name=""MeetingSuggestion"" type=""t:MeetingSuggestionType"" maxOccurs=""unbounded"" />
                                </xs:sequence>
                                </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.MeetingSuggestions.Length,
                    string.Format(
                    "There should be one meeting suggestion information in the entity extraction result, actual {0}.",
                    entityExtractionResult.MeetingSuggestions.Length));

                if (Common.IsRequirementEnabled(1276, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1276");

                    // The meetingSuggestion element is MeetingSuggestionType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        entityExtractionResult.MeetingSuggestions[0],
                        1276,
                        @"[In Appendix C: Product Behavior] Implementation does support the MeetingSuggestionType complex type which specifies a meeting suggestion. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""MeetingSuggestionType"">
                                    <xs:sequence>
                                      <xs:element name=""Attendees"" type=""t:ArrayOfEmailUsersType"" minOccurs=""0"" maxOccurs=""1"" />
                                      <xs:element name=""Location"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""Subject"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""MeetingString"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""StartTime"" type=""xs:dateTime"" minOccurs=""0"" />
                                      <xs:element name=""EndTime"" type=""xs:dateTime"" minOccurs=""0"" />
                                    </xs:sequence>
                                  </xs:complexType>");

                    this.VerifyEntityType(entityExtractionResult.MeetingSuggestions[0]);
                }
            }

            if (entityExtractionResult.TaskSuggestions != null)
            {
                if (Common.IsRequirementEnabled(1277, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1277");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1277
                    // The TaskSuggestions element is ArrayOfTaskSuggestionsType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1277,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfTaskSuggestionsType complex type which specifies an array of task suggestions.(Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfTaskSuggestionsType"">
                                <xs:sequence>
                                    <xs:element name=""TaskSuggestion"" type=""t:TaskSuggestionType"" maxOccurs=""unbounded"" />
                                </xs:sequence>
                                </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.TaskSuggestions.Length,
                    string.Format(
                    "There should be one task suggestion information in the entity extraction result, actual {0}.",
                    entityExtractionResult.TaskSuggestions.Length));

                this.VerifyTaskSuggestionType(entityExtractionResult.TaskSuggestions[0]);
            }

            if (entityExtractionResult.EmailAddresses != null)
            {
                if (Common.IsRequirementEnabled(1756, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1756");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1756
                    // The EmailAddresses element is ArrayOfEmailAddressEntitiesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1756,
                        @"[In Appendix C: Product Behavior] Implementation does use the ArrayOfEmailAddressEntitiesType complex type which specifies an array of email addresses. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfEmailAddressEntitiesType"">
                              <xs:sequence>
                                <xs:element name=""EmailAddressEntity"" type=""t:EmailAddressEntityType""
                                    minOccurs=""0"" maxOccurs=""unbounded""/>
                              </xs:sequence>
                            /xs:complexType>");
                }

                foreach (EmailAddressEntityType emailAddress in entityExtractionResult.EmailAddresses)
                {
                    if (Common.IsRequirementEnabled(1760, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1760");

                        // The emailAddress element is EmailAddressEntityType type, if schema is validated and the element is not null
                        // This requirement can be verified.
                        Site.CaptureRequirementIfIsNotNull(
                            emailAddress,
                            1760,
                            @"[In Appendix C: Product Behavior] Implementation does use this type [EmailAddressEntityType Complex Type] which extends the EntityType complex type, as specified in section 2.2.4.38. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""EmailAddressEntityType"">
                                    <xs:complexContent>
                                      <xs:extension base=""t:EntityType"">
                                        <xs:sequence>
                                          <xs:element name=""EmailAddress"" type=""xs:string"" minOccurs=""0""/>
                                        </xs:sequence>
                                      </xs:extension>
                                    </xs:complexContent>");

                        this.VerifyEntityType(emailAddress);
                    }
                }
            }

            if (entityExtractionResult.Contacts != null)
            {
                if (Common.IsRequirementEnabled(1278, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1278");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1278
                    // The Contacts element is ArrayOfContactsType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1278,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfContactsType complex type which specifies an array of contacts. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfContactsType"">
                                <xs:sequence>
                                  <xs:element name=""Contact"" type=""t:ContactType"" minOccurs=""0"" maxOccurs=""unbounded""/>
                                </xs:sequence>
                              </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.Contacts.Length,
                    string.Format(
                    "There should be one contact information in the entity extraction result, actual {0}.",
                    entityExtractionResult.Contacts.Length));

                this.VerifyContactType(entityExtractionResult.Contacts[0]);
            }

            if (entityExtractionResult.Urls != null)
            {
                if (Common.IsRequirementEnabled(1772, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1772");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1772
                    // The Urls element is ArrayOfUrlEntitiesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1772,
                        @"[In Appendix C: Product Behavior] Implementation does use the ArrayOfUrlEntitiesType complex type which specifies an array of URL entities. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfUrlEntitiesType"">
                              <xs:sequence>
                                <xs:element name=""UrlEntity"" type=""t:UrlEntityType"" minOccurs=""0"" maxOccurs=""unbounded""/>
                              </xs:sequence>
                            </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.Urls.Length,
                    string.Format(
                    "There should be one url information in the entity extraction result, actual {0}.",
                    entityExtractionResult.Urls.Length));

                if (Common.IsRequirementEnabled(1776, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1776");

                    // The url element is UrlEntityType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                        entityExtractionResult.Urls[0],
                        1776,
                        @"[In Appendix C: Product Behavior] Implementation does use this type [UrlEntityType Complex Type] which extends the EntityType complex type, as specified in section 2.2.4.38. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""UrlEntityType"">
                                  <xs:complexContent>
                                    <xs:extension base=""t:EntityType"">
                                      <xs:sequence>
                                        <xs:element name=""Url"" type=""xs:string"" minOccurs=""0""/>
                                      </xs:sequence>
                                    </xs:extension>
                                  </xs:complexContent>
                                </xs:complexType>");

                    this.VerifyEntityType(entityExtractionResult.Urls[0]);
                }
            }

            if (entityExtractionResult.PhoneNumbers != null)
            {
                if (Common.IsRequirementEnabled(1763, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1763");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1763
                    // The PhoneNumbers element is ArrayOfPhoneEntitiesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1763,
                        @"[In Appendix C: Product Behavior] Implementation does use the ArrayOfPhoneEntitiesType complex type which specifies an array of phone entities. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfPhoneEntitiesType"">
                              <xs:sequence>
                                <xs:element name=""Phone"" type=""t:PhoneEntityType"" minOccurs=""0"" maxOccurs=""unbounded""/>
                              </xs:sequence>
                            </xs:complexType>");
                }

                Site.Assert.AreEqual<int>(
                    1,
                    entityExtractionResult.PhoneNumbers.Length,
                    string.Format(
                    "There should be one phone number information in the entity extraction result, actual {0}.",
                    entityExtractionResult.PhoneNumbers.Length));

                if (Common.IsRequirementEnabled(1767, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1767");

                    // The phoneNumber element is PhoneEntityType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirementIfIsNotNull(
                                    entityExtractionResult.PhoneNumbers[0],
                                    1767,
                                    @"[In Appendix C: Product Behavior] Implementation does use this type [PhoneEntityType Complex Type] which extends the EntityType complex type, as specified in section 2.2.4.38. (Exchange 2013 and above follow this behavior.)
                                            <xs:complexType name=""PhoneEntityType"">
                                                <xs:complexContent>
                                                <xs:extension base=""t:EntityType"">
                                                    <xs:sequence>
                                                      <xs:element name=""OriginalPhoneString type=""xs:string"" minOccurs=""0""/>
                                                      <xs:element name=""PhoneString"" type=""xs:string"" minOccurs=""0""/>
                                                      <xs:element name=""Type"" type=""xs:string"" minOccurs=""0""/>
                                                    </xs:sequence>
                                                  </xs:extension>
                                                </xs:complexContent>
                                            </xs:complexType>");

                    this.VerifyEntityType(entityExtractionResult.PhoneNumbers[0]);
                }
            }
        }
        #endregion

        #region Verify EntityType Structure
        /// <summary>
        /// Verify the EntityType structure
        /// </summary>
        /// <param name="entity">An EntityType instance.</param>
        private void VerifyEntityType(EntityType entity)
        {
            if (Common.IsRequirementEnabled(1746, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1746");

                // The url element is UrlEntityType type, if schema is validated and the element is not null
                // This requirement can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    entity,
                    1746,
                    @"[In Appendix C: Product Behavior] Implementation does use the EntityType complex type in which entities are text found in certain parts of a message, such as the subject or the body. (Exchange 2013 and above follow this behavior.)
                        <xs:complexType name=""EntityType"">
                          <xs:sequence>
                            <xs:element name=""Position"" type=""t:EmailPositionType"" 
                                minOccurs=""0"" maxOccurs=""unbounded""/>
                          </xs:sequence>
                        </xs:complexType>");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1780");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1780
                // If the schema could be validated successfully, this requirement can be verified.
                Site.CaptureRequirement(
                    1780,
                    @"[In t:EmailPositionType Simple Type] [The type EmailPositionType is defined as follow:]
                        xs:simpleType name=""EmailPositionType"">
                          <xs:restriction base=""xs:string"">
                            <xs:enumeration value=""LatestReply""/>
                            <xs:enumeration value=""Other""/>
                            <xs:enumeration value=""Subject""/>
                            <xs:enumeration value=""Signature""/>
                          </xs:restriction>
                        </xs:simpleType>");
            }
        }
        #endregion

        #region Verify ContactType Structure
        /// <summary>
        /// Verify the ContactType structure
        /// </summary>
        /// <param name="contact">An ContactType instance.</param>
        private void VerifyContactType(ContactType contact)
        {
            if (Common.IsRequirementEnabled(1279, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1279");

                // The contact element is ContactType type, if schema is validated and the element is not null
                // This requirement can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    contact,
                    1279,
                    @"[In Appendix C: Product Behavior] Implementation does support the ContactType complex type which specifies the type of a contact. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""ContactType"">
                                    <xs:sequence>
                                      <xs:element name=""PersonName"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""BusinessName"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""PhoneNumbers"" type=""t:ArrayOfPhonesType"" minOccurs=""0"" maxOccurs=""1"" />
                                      <xs:element name=""Urls"" type=""t:ArrayOfUrlsType"" minOccurs=""0"" maxOccurs=""1"" />
                                      <xs:element name=""EmailAddresses"" type=""t:ArrayOfExtractedEmailAddresses"" minOccurs=""0"" maxOccurs=""1"" />
                                      <xs:element name=""Addresses"" type=""t:ArrayOfAddressesType"" minOccurs=""0"" maxOccurs=""1"" />
                                      <xs:element name=""ContactString"" type=""xs:string"" minOccurs=""0"" />
                                    </xs:sequence>
                                  </xs:complexType>");
            }

            if (contact.PhoneNumbers != null)
            {
                if (Common.IsRequirementEnabled(1281, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1281");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1281
                    // The PhoneNumbers element is ArrayOfPhonesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1281,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfPhonesType complex type which specifies an array of phone numbers. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfPhonesType"">
                                <xs:sequence>
                                    <xs:element name=""Phone"" type=""t:PhoneType"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                </xs:sequence>
                                </xs:complexType>");
                }

                foreach (PhoneType phoneNumber in contact.PhoneNumbers)
                {
                    if (Common.IsRequirementEnabled(1282, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1282");

                        // The phoneNumber element is PhoneEntityType type, if schema is validated and the element is not null
                        // This requirement can be verified.
                        Site.CaptureRequirementIfIsNotNull(
                            phoneNumber,
                            1282,
                            @"[In Appendix C: Product Behavior] Implementation does support the PhoneType complex type which specifies a phone number and its type. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""PhoneType"">
                                    <xs:sequence>
                                      <xs:element name=""OriginalPhoneString"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""PhoneString"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""Type"" type=""xs:string"" minOccurs=""0"" />
                                    </xs:sequence>
                                  </xs:complexType>");
                    }
                }
            }

            if (contact.Urls != null)
            {
                if (Common.IsRequirementEnabled(1280, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1280");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1280
                    // The Urls element is ArrayOfUrlsType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1280,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfUrlsType complex type which specifies an array of URLs. (Exchange 2013 and above follow this behavior.) 
                            <xs:complexType name=""ArrayOfUrlsType"">
                                <xs:sequence>
                                  <xs:element name=""Url"" type=""xs:string"" minOccurs=""0"" maxOccurs=""unbounded""/>
                                </xs:sequence>
                              </xs:complexType>");
                }
            }

            if (contact.EmailAddresses != null)
            {
                if (Common.IsRequirementEnabled(1287, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1287");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1287
                    // The EmailAddresses element is ArrayOfExtractedEmailAddresses type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1287,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfExtractedEmailAddresses complex type which specifies an array of email addresses. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfExtractedEmailAddresses"">
                                <xs:sequence>
                                  <xs:element name=""EmailAddress"" type=""xs:string"" minOccurs=""0"" maxOccurs=""unbounded""/>
                                </xs:sequence>
                              </xs:complexType>");
                }
            }

            if (contact.Addresses != null)
            {
                if (Common.IsRequirementEnabled(1272, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1272");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1272
                    // The Addresses element is ArrayOfAddressesType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1272,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfAddressesType complex type which specifies an array of addresses.(Exchange 2013 and above follow this behavior.)
                             <xs:complexType name=""ArrayOfAddressesType"">
                                <xs:sequence>
                                  <xs:element name=""Address"" type=""xs:string"" minOccurs=""0"" maxOccurs=""unbounded""/>
                                </xs:sequence>
                              </xs:complexType>");
                }
            }
        }
        #endregion

        #region Verify TaskSuggestionType Structure
        /// <summary>
        /// Verify the TaskSuggestionType structure
        /// </summary>
        /// <param name="taskSuggestion">An TaskSuggestionType instance.</param>
        private void VerifyTaskSuggestionType(TaskSuggestionType taskSuggestion)
        {
            if (Common.IsRequirementEnabled(1286, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1286");

                // The taskSuggestion element is TaskSuggestionType type, if schema is validated and the element is not null
                // This requirement can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    taskSuggestion,
                    1286,
                    @"[In Appendix C: Product Behavior] Implementation does support the TaskSuggestionType complex type which specifies a task suggestion. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""TaskSuggestionType"">
                                    <xs:sequence>
                                      <xs:element name=""TaskString"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""Assignees"" type=""t:EmailUserType"" minOccurs=""0"" maxOccurs=""1"" />
                                    </xs:sequence>
                                  </xs:complexType>");
            }

            if (taskSuggestion.Assignees != null)
            {
                if (Common.IsRequirementEnabled(1283, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1283");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1283
                    // The Assignees element is ArrayOfEmailUsersType type, if schema is validated and the element is not null
                    // This requirement can be verified.
                    Site.CaptureRequirement(
                        1283,
                        @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfEmailUsersType complex type which specifies an array of email users. (Exchange 2013 and above follow this behavior.)
                            <xs:complexType name=""ArrayOfEmailUsersType"">
                                <xs:sequence>
                                  <xs:element name=""EmailUser"" type=""t:EmailUserType"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                </xs:sequence>
                              </xs:complexType>");
                }

                foreach (EmailUserType assignee in taskSuggestion.Assignees)
                {
                    if (Common.IsRequirementEnabled(1284, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1284");

                        // The assignee element is EmailUserType type, if schema is validated and the element is not null
                        // This requirement can be verified.
                        Site.CaptureRequirementIfIsNotNull(
                            assignee,
                            1284,
                            @"[In Appendix C: Product Behavior] Implementation does support the EmailUserType complex type which specifies an email user. (Exchange 2013 and above follow this behavior.)
                                <xs:complexType name=""EmailUserType"">
                                    <xs:sequence>
                                      <xs:element name=""Name"" type=""xs:string"" minOccurs=""0"" />
                                      <xs:element name=""UserId"" type=""xs:string"" minOccurs=""0"" />
                                    </xs:sequence>
                                  </xs:complexType>");
                    }
                }
            }
        }
        #endregion

        #region Verify Structs Defined in MS-OXWSCDATA
        #region Verify ResponseMessageType Structure
        /// <summary>
        /// Verify the ResponseMessageType structure.
        /// </summary>
        /// <param name="responseMessage">A ResponseMessageType instance.</param>
        private void VerifyResponseMessageType(ResponseMessageType responseMessage)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1434");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1434
            Site.CaptureRequirementIfIsNotNull(
                responseMessage,
                "MS-OXWSCDATA",
                1434,
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

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1436");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1436
            bool isVerifiedR1436 = responseMessage.ResponseClass == ResponseClassType.Error ||
                responseMessage.ResponseClass == ResponseClassType.Success ||
                responseMessage.ResponseClass == ResponseClassType.Warning;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1436,
                "MS-OXWSCDATA",
                1436,
                @"[In m:ResponseMessageType Complex Type] [ResponseClass:] The following values are valid for this attribute: Success, Warning, Error.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1284");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1284
            Site.CaptureRequirementIfIsNotNull(
                responseMessage.ResponseClass,
                "MS-OXWSCDATA",
                1284,
                @"[In m:ResponseMessageType Complex Type] This attribute [ResponseClass] MUST be present.");

            // Verify the ResponseClassType schema.
            this.VerifyResponseClassType(responseMessage);
        }
        #endregion

        #region Verify ResponseClassType Structure
        /// <summary>
        /// Verify the ResponseClassType structure.
        /// </summary>
        /// <param name="responseMessage">A ResponseMessageType instance.</param>
        private void VerifyResponseClassType(ResponseMessageType responseMessage)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R191");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R191
            Site.CaptureRequirementIfIsNotNull(
                responseMessage.ResponseClass,
                "MS-OXWSCDATA",
                191,
                @"[In t:ResponseClassType Simple Type] The type [ResponseClassType] is defined as follow:
                    <xs:simpleType name=""ResponseClassType"">
                      <xs:restriction
                        base=""xs:string""
                      >
                        <xs:enumeration
                          value=""Error""
                         />
                        <xs:enumeration
                          value=""Success""
                         />
                        <xs:enumeration
                          value=""Warning""
                         />
                      </xs:restriction>
                    </xs:simpleType>");
        }
        #endregion

        #region Verify ArrayOfResponseMessagesType Structure
        /// <summary>
        /// Verify the ArrayOfResponseMessagesType structure.
        /// </summary>
        /// <param name="responseMessages">An ArrayOfResponseMessagesType instance.</param>
        private void VerifyArrayOfResponseMessagesType(ArrayOfResponseMessagesType responseMessages)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1036");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1036
            // The schema is validated and the responseMessages is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                responseMessages,
                "MS-OXWSCDATA",
                1036,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The type [ArrayOfResponseMessagesType] is defined as follow:
                <xs:complexType name=""ArrayOfResponseMessagesType"">
                  <xs:choice
                    maxOccurs=""unbounded""
                  >
                    <xs:element name=""CreateItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""DeleteItemResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""UpdateItemResponseMessage""
                      type=""m:UpdateItemResponseMessageType""
                     />
                    <xs:element name=""UpdateItemInRecoverableItemsResponseMessage"" 
                     type=""m:UpdateItemInRecoverableItemsResponseMessageType""/>
                    <xs:element name=""SendItemResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""DeleteFolderResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""EmptyFolderResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""CreateFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""GetFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""FindFolderResponseMessage""
                      type=""m:FindFolderResponseMessageType""
                     />
                    <xs:element name=""UpdateFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""MoveFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""CopyFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""CreateFolderPathResponseMessage"" 
                     type=""m:FolderInfoResponseMessageType""
                    />
                    <xs:element name=""CreateAttachmentResponseMessage""
                      type=""m:AttachmentInfoResponseMessageType""
                     />
                    <xs:element name=""DeleteAttachmentResponseMessage""
                      type=""m:DeleteAttachmentResponseMessageType""
                     />
                    <xs:element name=""GetAttachmentResponseMessage""
                      type=""m:AttachmentInfoResponseMessageType""
                     />
                    <xs:element name=""UploadItemsResponseMessage""
                      type=""m:UploadItemsResponseMessageType""
                     />
                    <xs:element name=""ExportItemsResponseMessage""
                      type=""m:ExportItemsResponseMessageType""
                     />
                    <xs:element name=""MarkAllItemsAsReadResponseMessage"" 
                       type=""m:ResponseMessageType""/>
                    <xs:element name=""GetClientAccessTokenResponseMessage"" 
                       type=""m:GetClientAccessTokenResponseMessageType""/>
                    <xs:element name=""GetAppManifestsResponseMessage"" type=""m:ResponseMessageType""/>
                    <xs:element name=""GetClientExtensionResponseMessage"" 
                       type=""m:ResponseMessageType""/>
                    <xs:element name=""SetClientExtensionResponseMessage"" 
                       type=""m:ResponseMessageType""/>

                    <xs:element name=""FindItemResponseMessage""
                      type=""m:FindItemResponseMessageType""
                     />
                    <xs:element name=""MoveItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""ArchiveItemResponseMessage"" type=""m:ItemInfoResponseMessageType""/>
                    <xs:element name=""CopyItemResponseMessage""
                      type=""m:ItemInfoResponseMessageType""
                     />
                    <xs:element name=""ResolveNamesResponseMessage""
                      type=""m:ResolveNamesResponseMessageType""
                     />
                    <xs:element name=""ExpandDLResponseMessage""
                      type=""m:ExpandDLResponseMessageType""
                     />
                    <xs:element name=""GetServerTimeZonesResponseMessage""
                      type=""m:GetServerTimeZonesResponseMessageType""
                     />
                    <xs:element name=""GetEventsResponseMessage""
                      type=""m:GetEventsResponseMessageType""
                     />
                    <xs:element name=""GetStreamingEventsResponseMessage""
                      type=""m:GetStreamingEventsResponseMessageType""
                     />
                    <xs:element name=""SubscribeResponseMessage""
                      type=""m:SubscribeResponseMessageType""
                     />
                    <xs:element name=""UnsubscribeResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""SendNotificationResponseMessage""
                      type=""m:SendNotificationResponseMessageType""
                     />
                    <xs:element name=""SyncFolderHierarchyResponseMessage""
                      type=""m:SyncFolderHierarchyResponseMessageType""
                     />
                    <xs:element name=""SyncFolderItemsResponseMessage""
                      type=""m:SyncFolderItemsResponseMessageType""
                     />
                    <xs:element name=""CreateManagedFolderResponseMessage""
                      type=""m:FolderInfoResponseMessageType""
                     />
                    <xs:element name=""ConvertIdResponseMessage""
                      type=""m:ConvertIdResponseMessageType""
                     />
                    <xs:element name=""GetSharingMetadataResponseMessage""
                      type=""m:GetSharingMetadataResponseMessageType""
                     />
                    <xs:element name=""RefreshSharingFolderResponseMessage""
                      type=""m:RefreshSharingFolderResponseMessageType""
                     />
                    <xs:element name=""GetSharingFolderResponseMessage""
                      type=""m:GetSharingFolderResponseMessageType""
                     />
                    <xs:element name=""CreateUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""DeleteUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetUserConfigurationResponseMessage""
                      type=""m:GetUserConfigurationResponseMessageType""
                     />
                    <xs:element name=""UpdateUserConfigurationResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""GetRoomListsResponse""
                      type=""m:GetRoomListsResponseMessageType""
                     />
                    <xs:element name=""GetRoomsResponse""
                      type=""m:GetRoomsResponseMessageType""
                     />
                      <xs:element name=""GetRemindersResponse"" 
                       type=""m:GetRemindersResponseMessageType""/>
                      <xs:element name=""PerformReminderActionResponse"" 
                       type=""m:PerformReminderActionResponseMessageType""/>
                    <xs:element name=""ApplyConversationActionResponseMessage""
                      type=""m:ResponseMessageType""
                     />
                    <xs:element name=""FindMailboxStatisticsByKeywordsResponseMessage"" type=""m:FindMailboxStatisticsByKeywordsResponseMessageType""/>
                    <xs:element name=""GetSearchableMailboxesResponseMessage"" type=""m:GetSearchableMailboxesResponseMessageType""/>
                    <xs:element name=""SearchMailboxesResponseMessage"" type=""m:SearchMailboxesResponseMessageType""/>
                    <xs:element name=""GetDiscoverySearchConfigurationResponseMessage"" type=""m:GetDiscoverySearchConfigurationResponseMessageType""/>
                    <xs:element name=""GetHoldOnMailboxesResponseMessage"" type=""m:GetHoldOnMailboxesResponseMessageType""/>
                    <xs:element name=""SetHoldOnMailboxesResponseMessage"" type=""m:SetHoldOnMailboxesResponseMessageType""/>
                      <xs:element name=""GetNonIndexableItemStatisticsResponseMessage"" type=""m:GetNonIndexableItemStatisticsResponseMessageType""/>
                      <!-- GetNonIndexableItemDetails response -->
                      <xs:element name=""GetNonIndexableItemDetailsResponseMessage"" type=""m:GetNonIndexableItemDetailsResponseMessageType""/>
                      <!-- GetUserHoldSettings response -->
                    <xs:element name=""FindPeopleResponseMessage"" type=""m:FindPeopleResponseMessageType""/>

                    <xs:element name=""GetPasswordExpirationDateResponse"" type=""m:GetPasswordExpirationDateResponseMessageType""
                    />
                      <xs:element name=""GetPersonaResponseMessage"" type=""m:GetPersonaResponseMessageType""/>
                      <xs:element name=""GetConversationItemsResponseMessage"" type=""m:GetConversationItemsResponseMessageType""/>
                      <xs:element name=""GetUserRetentionPolicyTagsResponseMessage"" type=""m:GetUserRetentionPolicyTagsResponseMessageType""/>
                      <xs:element name=""GetUserPhotoResponseMessage"" type=""m:GetUserPhotoResponseMessageType""/>
                      <xs:element name=""MarkAsJunkResponseMessage"" type=""m:MarkAsJunkResponseMessageType""/>
                  </xs:choice>
                </xs:complexType>
                ");
        }
        #endregion

        #region Verify ArrayOfRealItemsType Structure
        /// <summary>
        /// Verify the ArrayOfRealItemsType schema.
        /// </summary>
        /// <param name="arrayOfRealItem">An ArrayOfRealItemsType instance.</param>
        private void VerifyArrayOfRealItemsTypeSchema(ArrayOfRealItemsType arrayOfRealItem)
        {
            if (Common.IsRequirementEnabled(19240, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R19240");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R19240
                Site.CaptureRequirementIfIsNotNull(
                    arrayOfRealItem,
                    "MS-OXWSCDATA",
                    19240,
                    @"[In Appendix B: Product Behavior] Implementation does support ArrayOfRealItemsType Complex Type. (Exchange 2007 follows this behavior.)
                        <xs:complexType name=""ArrayOfRealItemsType"">
                         <xs:sequence>
                          <xs:choice
                           minOccurs=""0""
                           maxOccurs=""unbounded""
                          >
                           <xs:element name=""Item""
                            type=""t:ItemType""
                            />
                           <xs:element name=""Message""
                            type=""t:MessageType""
                            />
                           <xs:element name=""CalendarItem""
                            type=""t:CalendarItemType""
                            />
                           <xs:element name=""Contact""
                            type=""t:ContactItemType""
                            />
                           <xs:element name=""MeetingMessage""
                            type=""t:MeetingMessageType""
                            />
                           <xs:element name=""MeetingRequest""
                            type=""t:MeetingRequestMessageType""
                            />
                           <xs:element name=""MeetingResponse""
                            type=""t:MeetingResponseMessageType""
                            />
                           <xs:element name=""MeetingCancellation""
                            type=""t:MeetingCancellationMessageType""
                            />
                           <xs:element name=""Task""
                            type=""t:TaskType""
                            />
                           <xs:element name=""PostItem""
                            type=""t:PostItemType""
                            />
                          </xs:choice>
                         </xs:sequence>
                        </xs:complexType>");
            }

            if (Common.IsRequirementEnabled(19241, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R19241");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R19241
                Site.CaptureRequirementIfIsNotNull(
                    arrayOfRealItem,
                    "MS-OXWSCDATA",
                    19241,
                    @"[In Appendix B: Product Behavior] Implementation does support ArrayOfRealItemsType Complex Type. (Exchange 2010 and above follow this behavior.)
                        <xs:complexType name=""ArrayOfRealItemsType"">
                         <xs:sequence>
                          <xs:choice
                           minOccurs=""0""
                           maxOccurs=""unbounded""
                          >
                           <xs:element name=""Item""
                            type=""t:ItemType""
                            />
                           <xs:element name=""Message""
                            type=""t:MessageType""
                            />
                           <xs:element name=""CalendarItem""
                            type=""t:CalendarItemType""
                            />
                           <xs:element name=""Contact""
                            type=""t:ContactItemType""
                            />
                           <xs:element name=""DistributionList""
                            type=""t:DistributionListType""
                            />
                           <xs:element name=""MeetingMessage""
                            type=""t:MeetingMessageType""
                            />
                           <xs:element name=""MeetingRequest""
                            type=""t:MeetingRequestMessageType""
                            />
                           <xs:element name=""MeetingResponse""
                            type=""t:MeetingResponseMessageType""
                            />
                           <xs:element name=""MeetingCancellation""
                            type=""t:MeetingCancellationMessageType""
                            />
                           <xs:element name=""Task""
                            type=""t:TaskType""
                            />
                           <xs:element name=""PostItem""
                            type=""t:PostItemType""
                            />
                          </xs:choice>
                         </xs:sequence>
                        </xs:complexType>");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1675");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1675
            Site.CaptureRequirementIfIsNotNull(
                arrayOfRealItem,
                "MS-OXWSCDATA",
                1675,
                @"[In m:ItemInfoResponseMessageType Complex Type] The element ""Items"" is ""t:ArrayOfRealItemsType"" type (section 2.2.4.8).");
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
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1092");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1092
            Site.CaptureRequirementIfIsNotNull(
                baseResponseMessage,
                "MS-OXWSCDATA",
                1092,
                @"[In m:BaseResponseMessageType Complex Type] The type [BaseResponseMessageType] is defined as follow:
                <xs:complexType name=""BaseResponseMessageType"">
                  <xs:sequence>
                    <xs:element name=""ResponseMessages""
                      type=""m:ArrayOfResponseMessagesType""
                     />
                  </xs:sequence>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1623");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1623
            Site.CaptureRequirementIfIsNotNull(
                baseResponseMessage.ResponseMessages,
                "MS-OXWSCDATA",
                1623,
                @"[In m:BaseResponseMessageType Complex Type] The element ""ResponseMessages"" is ""m:ArrayOfResponseMessagesType""  type (section 2.2.4.10).");

            // Verify the ArrayOfResponseMessagesType schema.
            this.VerifyArrayOfResponseMessagesType(baseResponseMessage.ResponseMessages);

            // Because the requirement "MS-OXWSCDATA_R1092" contains the same information of this requirement, 
            // If the requirement "MS-OXWSCDATA_R1092" is verified, this requirement has been verified at the same time. 
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1094");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1094
            Site.CaptureRequirementIfIsNotNull(
                baseResponseMessage,
                "MS-OXWSCDATA",
                1094,
                @"[In m:BaseResponseMessageType Complex Type] There MUST be only one ResponseMessages element in a response.");

            foreach (ResponseMessageType message in baseResponseMessage.ResponseMessages.Items)
            {
                this.VerifyResponseMessageType(message);
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1085");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1085
            // BaseItemIdType is not used directly in the schema of each operations,
            // if schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1085,
                "[In t:BaseItemIdType Complex Type] The BaseItemIdType type MUST NOT be sent in a SOAP message because it is an abstract type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1091");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1091
            // BaseResponseMessageType is not used directly in the schema of each operations,
            // if schema is validated, this requirement can be captured directly.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1091,
                "[In m:BaseResponseMessageType Complex Type] The BaseResponseMessageType complex type MUST NOT be sent in a SOAP message because it is an abstract type.");
        }
        #endregion

        #region Verify ItemInfoResponseMessageType Structure
        /// <summary>
        /// Verify the ItemInfoResponseMessageType structure.
        /// </summary>
        /// <param name="itemInfoResponseMessage">An ItemInfoResponseMessageType instance.</param>
        private void VerifyItemInfoResponseMessageType(ItemInfoResponseMessageType itemInfoResponseMessage)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1181");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1181
            Site.CaptureRequirementIfIsNotNull(
                itemInfoResponseMessage,
                "MS-OXWSCDATA",
                1181,
                @"[In m:ItemInfoResponseMessageType Complex Type] The type [ItemInfoResponseMessageType] is defined as follow:
                    <xs:complexType name=""ItemInfoResponseMessageType"">
                      <xs:complexContent>
                        <xs:extension
                          base=""m:ResponseMessageType""
                        >
                          <xs:sequence>
                            <xs:element name=""Items""
                              type=""t:ArrayOfRealItemsType""
                             />
                          </xs:sequence>
                        </xs:extension>
                      </xs:complexContent>
                    </xs:complexType>");

            ArrayOfRealItemsType arrayOfRealItemsType = itemInfoResponseMessage.Items;

            // Verify the ArrayOfRealItemsType schema
            this.VerifyArrayOfRealItemsTypeSchema(arrayOfRealItemsType);

            // Verify the ItemType schema.
            if (itemInfoResponseMessage != null && itemInfoResponseMessage.ResponseClass == ResponseClassType.Success)
            {
                if (arrayOfRealItemsType != null && arrayOfRealItemsType.Items != null)
                {
                    foreach (ItemType item in arrayOfRealItemsType.Items)
                    {
                        this.VerifyItemType(item);
                    }
                }
            }
        }
        #endregion

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
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1339");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1339
            Site.CaptureRequirementIfIsNotNull(
                serverVersionInfo,
                "MS-OXWSCDATA",
                1339,
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
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1456");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1456
                // If MajorVersion element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1456,
                    "[In t:ServerVersionInfo Element] MajorVersion type is xs:int.");
            }

            if (serverVersionInfo.MinorVersionSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1457");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1457
                // If MinorVersion element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1457,
                    "[In t:ServerVersionInfo Element] MinorVersion type is xs:int.");
            }

            if (serverVersionInfo.MajorBuildNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1458");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1458
                // If MajorBuildNumber element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1458,
                    "[In t:ServerVersionInfo Element] MajorBuildNumber type is xs:int.");
            }

            if (serverVersionInfo.MinorBuildNumberSpecified)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1459");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1459
                // If MinorBuildNumber element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1459,
                    "[In t:ServerVersionInfo Element] MinorBuildNumber type is xs:int.");
            }

            if (serverVersionInfo.Version != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1460");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1460
                // If Version element is specified, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1460,
                    "[In t:ServerVersionInfo Element] Version type is xs:string.");
            }
        }
        #endregion
        #endregion

        #region Verify ItemIdType Structure
        /// <summary>
        /// Verify ItemIdType structure.
        /// </summary>
        /// <param name="response">A BaseResponseMessageType instance.</param>
        private void VerifyItemId(BaseResponseMessageType response)
        {
            ArrayOfResponseMessagesType responseMessages = response.ResponseMessages;
            foreach (ResponseMessageType responseMessage in responseMessages.Items)
            {
                if (responseMessage.ResponseCode != ResponseCodeType.NoError || responseMessage.ResponseClass != ResponseClassType.Success)
                {
                    continue;
                }

                ItemInfoResponseMessageType itemInfoResponseMessage = responseMessage as ItemInfoResponseMessageType;
                ArrayOfRealItemsType arrayOfRealItemsType = itemInfoResponseMessage.Items;
                if (arrayOfRealItemsType.Items == null)
                {
                    continue;
                }

                foreach (ItemType tempItem in arrayOfRealItemsType.Items)
                {
                    this.VerifyItemIdType(tempItem.ItemId);
                    if (tempItem.ConversationId != null)
                    {
                        this.VerifyItemIdType(tempItem.ConversationId);
                    }
                }
            }
        }
        #endregion

        #region Verify ItemId Defined in MS-OXWSITEMID
        /// <summary>
        /// Verify ItemId Defined in MS-OXWSITEMID.
        /// </summary>
        /// <param name="itemId">A ItemIdType instance.</param>
        private void VerifyItemIdType(ItemIdType itemId)
        {
            ItemIdId itemIdId = this.itemAdapter.ParseItemId(itemId);
            if (itemIdId.StorageType == IdStorageType.MailboxItemMailboxGuidBased || itemIdId.StorageType == IdStorageType.ConversationIdMailboxGuidBased)
            {
                bool isR37Verified = itemIdId.MonikerLength != 0 && itemIdId.MonikerGuid != null;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R37");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R37
                Site.CaptureRequirementIfIsTrue(
                    isR37Verified,
                    "MS-OXWSITEMID",
                    37,
                    @"[In MailboxItemMailboxGuidBased or ConversationIdMailboxGuidBased] Read the mailbox guid by doing the following.
                        Read Int16 from stream for the length.
                        Read 'length' number of bytes from the stream as byte[].
                        Return new Guid(Encoding.UTF8.GetString(moniker, 0, moniker.Length));");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R38");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R38
                // Id processing instrctuion is gotten in ParseItemId method, so this requirement can be captured directly.
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    38,
                    @"[In MailboxItemMailboxGuidBased or ConversationIdMailboxGuidBased] Read the Id processing instruction by doing the following.
                        Read byte from stream.
                        Cast value as IdProcessingInstruction enum value and return.");

                bool isR39Verified = itemIdId.StoreId != null && itemIdId.StoreIdLength != 0;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R39");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R39
                Site.CaptureRequirementIfIsTrue(
                    isR39Verified,
                    "MS-OXWSITEMID",
                    39,
                    @"[In MailboxItemMailboxGuidBased or ConversationIdMailboxGuidBased] Read store Id bytes (for conversationId or item id) by doing the following.
                        Read Int16 from stream for length.
                        Read 'length' number of bytes from stream.
                        Return as byte[].");

                if (itemIdId.StorageType == IdStorageType.MailboxItemMailboxGuidBased && Common.IsRequirementEnabled(74, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R74");

                    // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R74
                    // According above capture, the format of the ItemId has been verified to follow that is defined in R74.
                    // So if the Id storage type is MailboxItemMailboxGuidBased, then R74 has been verified.
                    Site.CaptureRequirement(
                        "MS-OXWSITEMID",
                        74,
                        @"[In Appendix A: Product Behavior] Implementation does support this value [MailboxItemMailboxGuidBased]. (Exchange 2007 Service Pack 1 (SP1) and above follow this behavior).");
                }

                if (itemIdId.StorageType == IdStorageType.ConversationIdMailboxGuidBased)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R75");

                    // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R75
                    // According above capture, the format of the ItemId has been verified to follow that is defined in R75.
                    // So if the Id storage type is ConversationIdMailboxGuidBased, then R75 has been verified.
                    Site.CaptureRequirement(
                        "MS-OXWSITEMID",
                        75,
                        @"[In MailboxItemMailboxGuidBased or ConversationIdMailboxGuidBased] If the Id storage type is ConversationIdMailboxGuidBased, the format of the remaining bytes is
                            [Short] Moniker Length
                            [Variable] Moniker Bytes
                            [Byte] Id Processing Instruction (Normal = 0, Recurrence = 1)
                            [Short] Store Id Bytes Length
                            [Variable] Store Id Bytes");
                }
            }

            if (itemIdId.StorageType == IdStorageType.PublicFolderItem)
            {
                bool isR44Verified = itemIdId.StoreId != null && itemIdId.StoreIdLength != 0 && itemIdId.FolderId != null && itemIdId.FolderIdLength != 0;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R44");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R44
                Site.CaptureRequirementIfIsTrue(
                    isR44Verified,
                    "MS-OXWSITEMID",
                    44,
                    @"[In PublicFolderItem] If the Id storage type is PublicFolderItem the format of the remaining bytes is:
                        [Byte] Id Processing Instruction
                        [Short] Store Id Bytes Length
                        [Variable] Store Id Bytes
                        [Short] Folder Id Bytes Length
                        [Variable] Folder Id Bytes
                        ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R46");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R46
                // Id processing instrctuion is gotten in ParseItemId method, so this requirement can be captured directly.
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    46,
                    @"[In PublicFolderItem] Read the Id processing instruction by doing the following:
                        Read byte from stream.
                        Cast value as IdProcessingInstruction enum value.");

                bool isR47Verified = itemIdId.StoreIdLength != 0;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R47");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R47
                Site.CaptureRequirementIfIsTrue(
                    isR47Verified,
                    "MS-OXWSITEMID",
                    47,
                    @"[In PublicFolderItem] Read the store Id bytes for the item Id by doing the following steps. 
                        Read Int16 from stream for length.
                        Read 'length' number of bytes from stream as byte[].");

                bool isR48Verified = itemIdId.FolderId != null;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R48");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R48
                Site.CaptureRequirementIfIsTrue(
                    isR48Verified,
                    "MS-OXWSITEMID",
                    48,
                    @"[In PublicFolderItem] Read the store Id bytes for parent folder Id by doing the following.
                        Read Int16 from stream for length.
                        Read 'length' number of bytes from stream as byte[].");
            }
        }
        #endregion
    } 
}