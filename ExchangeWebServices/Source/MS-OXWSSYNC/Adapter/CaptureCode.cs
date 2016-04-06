namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to verify data/operation format in MS-OXWSSYNC.
    /// </summary>
    public partial class MS_OXWSSYNCAdapter
    {
        /// <summary>
        /// Verify the SOAP version.
        /// </summary>
        private void VerifySoapVersion()
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R3");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R3
            // According to the implementation of adapter, the message is formatted as SOAP 1.1. If the operation is invoked successfully, then this requirement can be verified.
            Site.CaptureRequirement(
                3,
                @"[In Transport] The SOAP version supported is SOAP 1.1.");
        }

        /// <summary>
        /// Verify the transport related requirements.
        /// </summary>
        private void VerifyTransportType()
        {
            // Get the transport type
            TransportProtocol transport = (TransportProtocol)Enum.Parse(typeof(TransportProtocol), Common.GetConfigurationPropertyValue("TransportType", Site), true);

            if (transport == TransportProtocol.HTTPS)
            {
                if (Common.IsRequirementEnabled(20003, this.Site))
                {
                    // Add debug information.
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R20003");

                    // Verify MS-OXWSBTRF requirement: MS-OXWSSYNC_R20003
                    // Because Adapter uses SOAP and HTTPS to communicate with server, if server returned data without exception, this requirement has been captured.
                    Site.CaptureRequirement(
                        20003,
                        @"[In Appendix C: Product Behavior] Implementation does support SOAP over HTTPS, as specified in [RFC2818]. (Exchange 2007 and above follow this behavior.)");
                }
            }
            else if (transport == TransportProtocol.HTTP)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R20001");

                // Verify MS-OXWSBTRF requirement: MS-OXWSSYNC_R20001
                // Because Adapter uses SOAP and HTTP to communicate with server, if server returned data without exception, this requirement has been captured.
                Site.CaptureRequirement(
                    20001,
                    @"[In Transport]The protocol MUST support SOAP over HTTP, as specified in [RFC2616].");
            }
        }

        /// <summary>
        /// The capture code of requirements in SyncFolderHierarchy operation.
        /// </summary>
        /// <param name="syncFoldHierarchyResponse">The response message for SyncFolderHierarchy operation.</param>
        /// <param name="isSchemaValidated">A Boolean value indicates the schema validation result, true means the schema is validated, false means the schema is not validated.</param>
        private void VerifySyncFolderHierarchyResponse(SyncFolderHierarchyResponseType syncFoldHierarchyResponse, bool isSchemaValidated)
        {
            // Assert the response is not null
            Site.Assert.IsNotNull(syncFoldHierarchyResponse, "If the request is successful, the response should not be null.");

            // The SyncFolderHierarchy operation MUST return one SyncFolderHierarchyResponseMessage element.
            Site.Assert.AreEqual<int>(1, syncFoldHierarchyResponse.ResponseMessages.Items.Length, "The SyncFolderHierarchy operation MUST return one SyncFolderHierarchyResponseMessage element.");
            SyncFolderHierarchyResponseMessageType syncFoldHierarchyResponseMessage = (SyncFolderHierarchyResponseMessageType)syncFoldHierarchyResponse.ResponseMessages.Items[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R441");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R441
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                441,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] [The schema of ""SyncFolderHierarchyResponseMessageType"" is:]
                <xs:complexType name=""SyncFolderHierarchyResponseMessageType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""m:ResponseMessageType""
                    >
                      <xs:sequence>
                        <xs:element name=""SyncState""
                          type=""xs:string""
                          minOccurs=""0""
                         />
                        <xs:element name=""IncludesLastFolderInRange""
                          type=""xs:boolean""
                          minOccurs=""0""
                         />
                        <xs:element name=""Changes""
                          type=""t:SyncFolderHierarchyChangesType""
                          minOccurs=""0""
                         />
                      </xs:sequence>
                    </xs:extension>
                  </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R442");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R442
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                442,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] The type of SyncState is xs:string [XMLSCHEMA2]");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R443");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R443
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                443,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] The type of IncludesLastFolderInRange is xs:boolean [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R48");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R48.
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                48,
                @"[In m:SyncFolderHierarchyResponseMessageType Complex Type] The type of Changes is t:SyncFolderHierarchyChangesType (section 3.1.4.1.3.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R447");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R447
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                447,
                @"[In t:SyncFolderHierarchyChangesType Complex Type] The SyncFolderHierarchyChangesType complex type specifies a sequenced array of change types that describe the differences between the folders on the client and the folders on the server. 
                <xs:complexType name=""SyncFolderHierarchyChangesType"">
                  <xs:sequence>
                    <xs:choice
                      maxOccurs=""unbounded""
                      minOccurs=""0""
                    >
                      <xs:element name=""Create""
                        type=""t:SyncFolderHierarchyCreateOrUpdateType""
                       />
                      <xs:element name=""Update""
                        type=""t:SyncFolderHierarchyCreateOrUpdateType""
                       />
                      <xs:element name=""Delete""
                        type=""t:SyncFolderHierarchyDeleteType""
                       />
                    </xs:choice>
                  </xs:sequence>
                </xs:complexType>");

            if (null != syncFoldHierarchyResponseMessage.Changes && null != syncFoldHierarchyResponseMessage.Changes.Items)
            {
                for (int i = 0; i < syncFoldHierarchyResponseMessage.Changes.ItemsElementName.Length; i++)
                {
                    if (syncFoldHierarchyResponseMessage.Changes.ItemsElementName[i] == ItemsChoiceType.Create)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R83");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R83
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            83,
                            @"[In t:SyncFolderHierarchyChangesType Complex Type] The type of Create is t:SyncFolderHierarchyCreateOrUpdateType (section 3.1.4.1.3.2).");
                    }

                    // If the change is updated, verify MS-OXWSSYNC_R448 and requirements in update operation.
                    if (syncFoldHierarchyResponseMessage.Changes.ItemsElementName[i] == ItemsChoiceType.Update)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R448");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R448
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            448,
                            @"[In t:SyncFolderHierarchyChangesType Complex Type] The type of Update is t:SyncFolderHierarchyCreateOrUpdateType.");
                    }

                    // If the change is updated or created, verify MS-OXWSSYNC_R449 and requirements in update operation or create operation.
                    if (syncFoldHierarchyResponseMessage.Changes.ItemsElementName[i] == ItemsChoiceType.Update || syncFoldHierarchyResponseMessage.Changes.ItemsElementName[i] == ItemsChoiceType.Create)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R449");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R449
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            449,
                            @"[In t:SyncFolderHierarchyCreateOrUpdateType Complex Type] The SyncFolderHierarchyCreateOrUpdateType complex type specifies a single folder to create or update in the client data store. 
xs:complexType name=""SyncFolderHierarchyCreateOrUpdateType"">
  <xs:choice>
    <xs:element name=""Folder""
      type=""t:FolderType""
     />
    <xs:element name=""CalendarFolder""
      type=""t:CalendarFolderType""
     />
    <xs:element name=""ContactsFolder""
      type=""t:ContactsFolderType""
     />
    <xs:element name=""SearchFolder""
      type=""t:SearchFolderType""
     />
    <xs:element name=""TasksFolder""
      type=""t:TasksFolderType""
     />
  </xs:choice>
</xs:complexType>
");
                    }

                    // If the change is deleted, verify MS-OXWSSYNC_R86 and requirements in delete operation.
                    if (syncFoldHierarchyResponseMessage.Changes.ItemsElementName[i] == ItemsChoiceType.Delete)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R86");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R86
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            86,
                            @"[In t:SyncFolderHierarchyChangesType Complex Type] The type of Delete is  t:SyncFolderHierarchyDeleteType (section 3.1.4.1.3.3).");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R450");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R450
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            450,
                            @"[In t:SyncFolderHierarchyDeleteType Complex Type] The SyncFolderHierarchyDeleteType complex type specifies a folder to delete from the client data store. 
                        <xs:complexType name=""SyncFolderHierarchyDeleteType"">
                          <xs:sequence>
                            <xs:element name=""FolderId""
                              type=""t:FolderIdType""
                             />
                          </xs:sequence>
                        </xs:complexType>");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1165");

                        // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1165
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
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

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R116");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R116
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            116,
                            @"[In t:SyncFolderHierarchyDeleteType Complex Type] The type of FolderId is t:FolderIdType ([MS-OXWSCDATA] section 2.2.4.36).");
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R457");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R457
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                457,
                @"[In SyncFolderHierarchy] The following is the WSDL port type specification of the SyncFolderHierarchy operation.
                <wsdl:operation name=""SyncFolderHierarchy"">
                     <wsdl:input message=""tns:SyncFolderHierarchySoapIn"" />
                     <wsdl:output message=""tns:SyncFolderHierarchySoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R459");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R459
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                459,
                @"[In SyncFolderHierarchy] The following is the WSDL binding specification of the SyncFolderHierarchy operation.
                    <wsdl:operation name=""SyncFolderHierarchy"">
                       <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/SyncFolderHierarchy"" />
                       <wsdl:input>
                          <soap:header message=""tns:SyncFolderHierarchySoapIn"" part=""Impersonation"" use=""literal""></soap:header>
                          <soap:header message=""tns:SyncFolderHierarchySoapIn"" part=""MailboxCulture"" use=""literal""/>
                          <soap:header message=""tns:SyncFolderHierarchySoapIn"" part=""RequestVersion"" use=""literal""/>
                          <soap:body parts=""request"" use=""literal"" />
                       </wsdl:input>
                       <wsdl:output>
                          <soap:body parts=""SyncFolderHierarchyResult"" use=""literal"" />
                          <soap:header message=""tns:SyncFolderHierarchySoapOut"" part=""ServerVersion"" use=""literal""/>
                       </wsdl:output>
                    </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R460");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R460
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                460,
                @"[In SyncFolderHierarchy] The SyncFolderHierarchy operation MUST return one SyncFolderHierarchyResponseMessage element in the ResponseMessages element of the SyncFolderHierarchyResponse element (section 3.1.4.1.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R463");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R463
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                463,
                @"[In m:SyncFolderHierarchyResponseType Complex Type] [The schema of ""SyncFolderHierarchyResponseType"" is:]
                <xs:complexType name=""SyncFolderHierarchyResponseType"">
                    <xs:complexContent>
                    <xs:extension
                        base=""m:BaseResponseMessageType""
                        />
                    </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R465");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R465
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                465,
                @"[In SyncFolderHierarchyResponse Element] The SyncFolderHierarchyResponse element specifies the response message for a SyncFolderHierarchy operation (section 3.1.4.1).  <xs:element name=""SyncFolderHierarchyResponse""
                type=""m:SyncFolderHierarchyResponseType""
                />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R299");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R299
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                299,
                @"[In tns:SyncFolderHierarchySoapOut] The SyncFolderHiearchySoapOut WSDL message specifies the server responses to a SyncFolderHierarchy operation request to return synchronization information.
                <wsdl:message name=""SyncFolderHierarchySoapOut"">
                <wsdl:part name=""SyncFolderHierarchyResult"" element=""tns:SyncFolderHierarchyResponse"" />
                <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R306");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R306
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                306,
                @"[In tns:SyncFolderHierarchySoapOut] The Element/Type of SyncFolderHierarchyResult is tns:SyncFolderHierarchyResponse (section 3.1.4.1.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R307");

            // Verify requirement MS-OXWSSYNC_R307
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                307,
                @"[In tns:SyncFolderHierarchySoapOut] [The part name syncFolderHierarchyResult] specifies the SOAP body of the response to a SyncFolderHiearchy operation request.");

            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R308");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R308
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    308,
                    @"[In tns:SyncFolderHierarchySoapOut] The Element/Type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R309");

                // Verify requirement MS-OXWSSYNC_R309
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    309,
                    @"[In tns:SyncFolderHierarchySoapOut] [The part name ServerVersion] specifies a SOAP header that identifies the server version for the response to a SyncFolderHiearchy operation request.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R238");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R238
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                238,
                @"[In SyncFolderHierarchy] The ResponseMessages element is specified as an element of the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.18).");

            // Verify the requirements in MS-OXWSCDATA.
            this.VerifyRequirementOfOXWSCDATA(syncFoldHierarchyResponseMessage, isSchemaValidated);
        }

        /// <summary>
        /// The capture code of requirements in SyncFolderItems operation.
        /// </summary>
        /// <param name="syncFolderItemsResponse">The response message for SyncFolderItems operation</param>
        /// <param name="isSchemaValidated">A Boolean value indicates the schema validation result, true means the schema is validated, false means the schema is not validated.</param>
        private void VerifySyncFolderItemsResponse(SyncFolderItemsResponseType syncFolderItemsResponse, bool isSchemaValidated)
        {
            // Assert the response is not null
            Site.Assert.IsNotNull(syncFolderItemsResponse, "If the request is successful, the response should not be null.");

            // If only one item is created in test case. The count of items in SyncFolderItems operation response should be one.
            Site.Assert.AreEqual<int>(1, syncFolderItemsResponse.ResponseMessages.Items.Length, "The count of items in SyncFolderItems operation response should be one if only one item is created.");
            SyncFolderItemsResponseMessageType syncFolderItemsResponseMessage = (SyncFolderItemsResponseMessageType)syncFolderItemsResponse.ResponseMessages.Items[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R444");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R444 
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                444,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] [The schema of ""SyncFolderItemsResponseMessageType"" is:]
                <xs:complexType name=""SyncFolderItemsResponseMessageType"">
                    <xs:complexContent>
                    <xs:extension
                        base=""m:ResponseMessageType""
                    >
                        <xs:sequence>
                        <xs:element name=""SyncState""
                            type=""xs:string""
                            minOccurs=""0""
                            />
                        <xs:element name=""IncludesLastItemInRange""
                            type=""xs:boolean""
                            minOccurs=""0""
                            />
                        <xs:element name=""Changes""
                            type=""t:SyncFolderItemsChangesType""
                            minOccurs=""0""
                            />
                        </xs:sequence>
                    </xs:extension>
                    </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R51");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R51 
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                51,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] The SyncFolderItemsResponseMessageType complex type specifies the status and results of a single call to the SyncFolderItems operation (section 3.1.4.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R445");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R445
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                445,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] The type of SyncState is xs:string ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R446");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R446
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                446,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] The type of IncludesLastItemInRange is xs:boolean ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R69");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R69
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                69,
                @"[In m:SyncFolderItemsResponseMessageType Complex Type] The type of Changes is t:SyncFolderItemsChangesType (section 2.2.4.5).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R451");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R451
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                451,
                @"[In t:SyncFolderItemsChangesType Complex Type] The SyncFolderItemsChangesType complex type specifies a sequenced array of change types that describe the differences between the items on the client and the items on the server.
                <xs:complexType name=""SyncFolderItemsChangesType"">
                    <xs:sequence>
                    <xs:choice
                        maxOccurs=""unbounded""
                        minOccurs=""0""
                    >
                        <xs:element name=""Create""
                        type=""t:SyncFolderItemsCreateOrUpdateType""
                        />
                        <xs:element name=""Update""
                        type=""t:SyncFolderItemsCreateOrUpdateType""
                        />
                        <xs:element name=""Delete""
                        type=""t:SyncFolderItemsDeleteType""
                        />
                        <xs:element name=""ReadFlagChange""
                        type=""t:SyncFolderItemsReadFlagType""
                        />
                    </xs:choice>
                    </xs:sequence>
                </xs:complexType>");

            if (null != syncFolderItemsResponseMessage.Changes && null != syncFolderItemsResponseMessage.Changes.Items)
            {
                // Verify the type of Create or Update.
                foreach (ItemsChoiceType1 itemsElementName in syncFolderItemsResponseMessage.Changes.ItemsElementName)
                {
                    if (itemsElementName == ItemsChoiceType1.Create)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R130");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R130
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            130,
                            @"[In t:SyncFolderItemsChangesType Complex Type] The type of Create is t:SyncFolderItemsCreateOrUpdateType (section 3.1.4.2.3.3).");
                    }
                    else if (itemsElementName == ItemsChoiceType1.Update)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R452");

                        // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R452
                        Site.CaptureRequirementIfIsTrue(
                            isSchemaValidated,
                            452,
                            @"[In t:SyncFolderItemsChangesType Complex Type] The type of Update is t:SyncFolderItemsCreateOrUpdateType.");
                    }

                    foreach (object item in syncFolderItemsResponseMessage.Changes.Items)
                    {
                        if (item.GetType() == typeof(SyncFolderItemsCreateOrUpdateType))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R453");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R453
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                453,
                                @"[In t:SyncFolderItemsCreateOrUpdateType Complex Type] The SyncFolderItemsCreateOrUpdateType complex type specifies a single item to create or update in the client data store. 
                        <xs:complexType name=""SyncFolderItemsCreateOrUpdateType"">
                            <xs:choice>
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
                            <xs:element name=""RoleMember"" 
                                type=""t:RoleMemberItemType""
                                />
                            <xs:element name=""Network"" 
                                type=""t:NetworkItemType""
                                />
                            <xs:element name=""Person""
                                 type=""t:AbchPersonItemType""
                                />
                            <xs:element name=""Booking""
                               type=""t:BookingItemType""
                                />
                            </xs:choice>
                        </xs:complexType>");
                        }
                        else if (item.GetType() == typeof(SyncFolderItemsDeleteType))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R133");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R133
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                133,
                                @"[In t:SyncFolderItemsChangesType Complex Type] The type of Delete is t:SyncFolderItemsDeleteType (section 3.1.4.2.3.4).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R183");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R183
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                183,
                                @"[In t:SyncFolderItemsDeleteType Complex Type] The type of ItemId is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R454");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R454
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                454,
                                @"[In t:SyncFolderItemsDeleteType Complex Type] The SyncFolderItemsDeleteType complex type specifies an item to delete from the client message store.
                        <xs:complexType name=""SyncFolderItemsDeleteType"">
                          <xs:sequence>
                            <xs:element name=""ItemId""
                              type=""t:ItemIdType""
                             />
                          </xs:sequence>
                        </xs:complexType>");
                        }
                        else if (item.GetType() == typeof(SyncFolderItemsReadFlagType))
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R135");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R135
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                135,
                                @"[In t:SyncFolderItemsChangesType Complex Type] The type of ReadFlagChange is t:SyncFolderItemsReadFlagType (section 3.1.4.2.3.5).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R455");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R455
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                455,
                                @"[In t:SyncFolderItemsReadFlagType Complex Type] The SyncFolderItemsReadFlagType complex type specifies whether an item on the server has been read. 
                        <xs:complexType name=""SyncFolderItemsReadFlagType"">
                            <xs:sequence>
                            <xs:element name=""ItemID""
                                type=""t:ItemIdType""
                                />
                            <xs:element name=""IsRead""
                                type=""xs:boolean""
                                />
                            </xs:sequence>
                        </xs:complexType>");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R456");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R456
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                456,
                                @"[In t:SyncFolderItemsReadFlagType Complex Type] The type of  IsRead is xs:boolean ([XMLSCHEMA2]).");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R193");

                            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R193
                            Site.CaptureRequirementIfIsTrue(
                                isSchemaValidated,
                                193,
                                @"[In t:SyncFolderItemsReadFlagType Complex Type] The type of ItemID is t:ItemIdType ([MS-OXWSCORE] section 2.2.4.25).");
                        }
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R468");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R468
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                468,
                @"[In SyncFolderItems] The following is the WSDL port type specification of the SyncFolderItems operation.
                <wsdl:operation name=""SyncFolderItems"">
                     <wsdl:input message=""tns:SyncFolderItemsSoapIn"" />
                     <wsdl:output message=""tns:SyncFolderItemsSoapOut"" />
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R469");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R469
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                469,
                @"[In SyncFolderItems] The following is the WSDL binding specification of the SyncFolderItems operation.
                <wsdl:operation name=""SyncFolderItems"">
                   <soap:operation soapAction=""http://schemas.microsoft.com/exchange/services/2006/messages/SyncFolderItems"" />
                   <wsdl:input>
                      <soap:header message=""tns:SyncFolderItemsSoapIn"" part=""Impersonation"" use=""literal""></soap:header>
                      <soap:header message=""tns:SyncFolderItemsSoapIn"" part=""MailboxCulture"" use=""literal""/>
                      <soap:header message=""tns:SyncFolderItemsSoapIn"" part=""RequestVersion"" use=""literal""/>
                      <soap:body parts=""request"" use=""literal"" />
                   </wsdl:input>
                   <wsdl:output>
                      <soap:body parts=""SyncFolderItemsResult"" use=""literal"" />
                      <soap:header message=""tns:SyncFolderItemsSoapOut"" part=""ServerVersion"" use=""literal""/>
                   </wsdl:output>
                </wsdl:operation>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R473");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R473
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                473,
                @"[In m:SyncFolderItemsResponseType Complex Type] [The schema of ""SyncFolderItemsResponseType"" is:]
                <xs:complexType name=""SyncFolderItemsResponseType"">
                  <xs:complexContent>
                    <xs:extension
                      base=""m:BaseResponseMessageType""
                     />
                  </xs:complexContent>
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R477");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R477
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                477,
                @"[In SyncFolderItemsResponse Element] The SyncFolderItemsResponse element specifies the response message for the SyncFolderItems operation. <xs:element name=""SyncFolderItemsResponse""
                  type=""m:SyncFolderItemsResponseType""
                 />");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R480");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R480
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                480,
                @"[In tns:SyncFolderItemsSoapOut] The SyncFolderItemsSoapOut WSDL message specifies the response from the SyncFoldersItems operation.
                <wsdl:message name=""SyncFolderItemsSoapOut"">
                   <wsdl:part name=""SyncFolderItemsResult"" element=""tns:SyncFolderItemsResponse"" />
                   <wsdl:part name=""ServerVersion"" element=""t:ServerVersionInfo""/>
                </wsdl:message>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R437");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R437
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                437,
                @"[In tns:SyncFolderItemsSoapOut] The Element/type of SyncFolderItemsResult is tns:SyncFolderItemsResponse (section 3.1.4.2.2.2).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R438");

            // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R438
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                438,
                @"[In tns:SyncFolderItemsSoapOut] [The part name SyncFolderItemsResult] specifies the SOAP body of the response to a SyncFolderItems operation request.");

            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R439");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R439
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    439,
                    @"[In tns:SyncFolderItemsSoapOut] The Element/type of ServerVersion is t:ServerVersionInfo ([MS-OXWSCDATA] section 2.2.3.12).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSSYNC_R440");

                // Verify MS-OXWSSYNC requirement: MS-OXWSSYNC_R440
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    440,
                    @"[In tns:SyncFolderItemsSoapOut] [The part name ServerVersion] specifies a SOAP header that identifies the server version for the response.");
            }

            // Verify the requirements in MS-OXWSCDATA.
            this.VerifyRequirementOfOXWSCDATA(syncFolderItemsResponseMessage, isSchemaValidated);
        }

        /// <summary>
        /// The capture code of requirements in MS-OXWSCDATA.
        /// </summary>
        /// <param name="responseMessage">ResponseMessageType responseMessage</param>
        /// <param name="isSchemaValidated">A Boolean value indicates the schema validation result, true means the schema is validated, false means the schema is not validated.</param>
        private void VerifyRequirementOfOXWSCDATA(ResponseMessageType responseMessage, bool isSchemaValidated)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1091");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1091
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1091,
                @"[In m:BaseResponseMessageType Complex Type] The BaseResponseMessageType complex type MUST NOT be sent in a SOAP message because it is an abstract type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1092");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1092
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
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

            // Add debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1036");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1036
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
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
                </xs:complexType>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1094");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1094
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
                "MS-OXWSCDATA",
                1094,
                @"[In m:BaseResponseMessageType Complex Type] There MUST be only one ResponseMessages element in a response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1434");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1434
            Site.CaptureRequirementIfIsTrue(
                isSchemaValidated,
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
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1436. Expected response class: 'Success', or 'Warning', or 'Error', actual response class: {0}", responseMessage.ResponseClass);

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1436
            bool isVerifyR1436 = (responseMessage.ResponseClass == ResponseClassType.Success) || (responseMessage.ResponseClass == ResponseClassType.Error) || (responseMessage.ResponseClass == ResponseClassType.Warning);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1436,
                "MS-OXWSCDATA",
                1436,
                @"[In m:ResponseMessageType Complex Type] [ResponseClass:] The following values are valid for this attribute: 
                Success,
                Warning,
                Error.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1284");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1284
            Site.CaptureRequirementIfIsNotNull(
                responseMessage.ResponseClass,
                "MS-OXWSCDATA",
                1284,
                @"[In m:ResponseMessageType Complex Type] This attribute [ResponseClass] MUST be present.");

            if (responseMessage.ResponseCodeSpecified == true)
            {
                // Add debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R197");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R197
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
                    "MS-OXWSCDATA",
                    197,
                    @"[In m:ResponseCodeType Simple Type] The type [ResponseCodeType] is defined as follow:
                <xs:simpleType name=""ResponseCodeType"">
                    <xs:restriction base=""xs:string"">
                        <xs:enumeration value=""NoError""/>
                        <xs:enumeration value=""ErrorAccessDenied""/>
                        <xs:enumeration value=""ErrorAccessModeSpecified""/>
                        <xs:enumeration value=""ErrorAccountDisabled""/>
                        <xs:enumeration value=""ErrorAddDelegatesFailed""/>
                        <xs:enumeration value=""ErrorAddressSpaceNotFound""/>
                        <xs:enumeration value=""ErrorADOperation""/>
                        <xs:enumeration value=""ErrorADSessionFilter""/>
                        <xs:enumeration value=""ErrorADUnavailable""/>
                        <xs:enumeration value=""ErrorAffectedTaskOccurrencesRequired""/>
                        <xs:enumeration value=""ErrorArchiveFolderPathCreation""/>
                        <xs:enumeration value=""ErrorArchiveMailboxNotEnabled""/>
                        <xs:enumeration value=""ErrorArchiveMailboxServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorAvailabilityConfigNotFound""/>
                        <xs:enumeration value=""ErrorBatchProcessingStopped""/>
                        <xs:enumeration value=""ErrorCalendarCannotMoveOrCopyOccurrence""/>
                        <xs:enumeration value=""ErrorCalendarCannotUpdateDeletedItem""/>
                        <xs:enumeration value=""ErrorCalendarCannotUseIdForOccurrenceId""/>
                        <xs:enumeration value=""ErrorCalendarCannotUseIdForRecurringMasterId""/>
                        <xs:enumeration value=""ErrorCalendarDurationIsTooLong""/>
                        <xs:enumeration value=""ErrorCalendarEndDateIsEarlierThanStartDate""/>
                        <xs:enumeration value=""ErrorCalendarFolderIsInvalidForCalendarView""/>
                        <xs:enumeration value=""ErrorCalendarInvalidAttributeValue""/>
                        <xs:enumeration value=""ErrorCalendarInvalidDayForTimeChangePattern""/>
                        <xs:enumeration value=""ErrorCalendarInvalidDayForWeeklyRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarInvalidPropertyState""/>
                        <xs:enumeration value=""ErrorCalendarInvalidPropertyValue""/>
                        <xs:enumeration value=""ErrorCalendarInvalidRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarInvalidTimeZone""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsCancelledForTentative""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsDelegatedForTentative""/>
                        <xs:enumeration value=""ErrorCalendarIsNotOrganizer""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForAccept""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForDecline""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForRemove""/>
                        <xs:enumeration value=""ErrorCalendarIsOrganizerForTentative""/>
                        <xs:enumeration
                             value=""ErrorCalendarOccurrenceIndexIsOutOfRecurrenceRange""/>
                        <xs:enumeration value=""ErrorCalendarOccurrenceIsDeletedFromRecurrence""/>
                        <xs:enumeration value=""ErrorCalendarOutOfRange""/>
                        <xs:enumeration value=""ErrorCalendarMeetingRequestIsOutOfDate""/>
                        <xs:enumeration value=""ErrorCalendarViewRangeTooBig""/>
                        <xs:enumeration value=""ErrorCallerIsInvalidADAccount""/>
                        <xs:enumeration value=""ErrorCannotArchiveCalendarContactTaskFolderException""/>
                        <xs:enumeration value=""ErrorCannotArchiveItemsInPublicFolders""/>
                        <xs:enumeration value=""ErrorCannotArchiveItemsInArchiveMailbo""/>
                        <xs:enumeration value=""ErrorCannotCreateCalendarItemInNonCalendarFolder""/>
                        <xs:enumeration value=""ErrorCannotCreateContactInNonContactFolder""/>
                        <xs:enumeration value=""ErrorCannotCreatePostItemInNonMailFolder""/>
                        <xs:enumeration value=""ErrorCannotCreateTaskInNonTaskFolder""/>
                        <xs:enumeration value=""ErrorCannotDeleteObject""/>
                        <xs:enumeration value=""ErrorCannotDisableMandatoryExtension""/>
                        <xs:enumeration value=""ErrorCannotGetSourceFolderPath""/>
                        <xs:enumeration value=""ErrorCannotGetExternalEcpUrl""/>
                        <xs:enumeration value=""ErrorCannotOpenFileAttachment""/>
                        <xs:enumeration value=""ErrorCannotDeleteTaskOccurrence""/>
                        <xs:enumeration value=""ErrorCannotEmptyFolder""/>
                        <xs:enumeration 
                            value=""ErrorCannotSetCalendarPermissionOnNonCalendarFolder""/>
                        <xs:enumeration 
                            value=""ErrorCannotSetNonCalendarPermissionOnCalendarFolder""/>
                        <xs:enumeration value=""ErrorCannotSetPermissionUnknownEntries""/>
                        <xs:enumeration value=""ErrorCannotSpecifySearchFolderAsSourceFolder""/>
                        <xs:enumeration value=""ErrorCannotUseFolderIdForItemId""/>
                        <xs:enumeration value=""ErrorCannotUseItemIdForFolderId""/>
                        <xs:enumeration value=""ErrorChangeKeyRequired""/>
                        <xs:enumeration value=""ErrorChangeKeyRequiredForWriteOperations""/>
                        <xs:enumeration value=""ErrorClientDisconnected""/>
                        <xs:enumeration value=""ErrorClientIntentInvalidStateDefinition""/>
                        <xs:enumeration value=""ErrorClientIntentNotFound""/>
                        <xs:enumeration value=""ErrorConnectionFailed""/>
                        <xs:enumeration value=""ErrorContainsFilterWrongType""/>
                        <xs:enumeration value=""ErrorContentConversionFailed""/>
                        <xs:enumeration value=""ErrorContentIndexingNotEnabled""/>
                        <xs:enumeration value=""ErrorCorruptData""/>
                        <xs:enumeration value=""ErrorCreateItemAccessDenied""/>
                        <xs:enumeration value=""ErrorCreateManagedFolderPartialCompletion""/>
                        <xs:enumeration value=""ErrorCreateSubfolderAccessDenied""/>
                        <xs:enumeration value=""ErrorCrossMailboxMoveCopy""/>
                        <xs:enumeration value=""ErrorCrossSiteRequest""/>
                        <xs:enumeration value=""ErrorDataSizeLimitExceeded""/>
                        <xs:enumeration value=""ErrorDataSourceOperation""/>
                        <xs:enumeration value=""ErrorDelegateAlreadyExists""/>
                        <xs:enumeration value=""ErrorDelegateCannotAddOwner""/>
                        <xs:enumeration value=""ErrorDelegateMissingConfiguration""/>
                        <xs:enumeration value=""ErrorDelegateNoUser""/>
                        <xs:enumeration value=""ErrorDelegateValidationFailed""/>
                        <xs:enumeration value=""ErrorDeleteDistinguishedFolder""/>
                        <xs:enumeration value=""ErrorDeleteItemsFailed""/>
                        <xs:enumeration value=""ErrorDeleteUnifiedMessagingPromptFailed""/>
                        <xs:enumeration value=""ErrorDistinguishedUserNotSupported""/>
                        <xs:enumeration value=""ErrorDistributionListMemberNotExist""/>
                        <xs:enumeration value=""ErrorDuplicateInputFolderNames""/>
                        <xs:enumeration value=""ErrorDuplicateUserIdsSpecified""/>
                        <xs:enumeration value=""ErrorEmailAddressMismatch""/>
                        <xs:enumeration value=""ErrorEventNotFound""/>
                        <xs:enumeration value=""ErrorExceededConnectionCount""/>
                        <xs:enumeration value=""ErrorExceededSubscriptionCount""/>
                        <xs:enumeration value=""ErrorExceededFindCountLimit""/>
                        <xs:enumeration value=""ErrorExpiredSubscription""/>
                        <xs:enumeration value=""ErrorExtensionNotFound""/>
                        <xs:enumeration value=""ErrorFolderCorrupt""/>
                        <xs:enumeration value=""ErrorFolderNotFound""/>
                        <xs:enumeration value=""ErrorFolderPropertRequestFailed""/>
                        <xs:enumeration value=""ErrorFolderSave""/>
                        <xs:enumeration value=""ErrorFolderSaveFailed""/>
                        <xs:enumeration value=""ErrorFolderSavePropertyError""/>
                        <xs:enumeration value=""ErrorFolderExists""/>
                        <xs:enumeration value=""ErrorFreeBusyGenerationFailed""/>
                        <xs:enumeration value=""ErrorGetServerSecurityDescriptorFailed""/>
                        <xs:enumeration value=""ErrorImContactLimitReached""/>
                        <xs:enumeration value=""ErrorImGroupDisplayNameAlreadyExists""/>
                        <xs:enumeration value=""ErrorImGroupLimitReached""/>
                        <xs:enumeration value=""ErrorImpersonateUserDenied""/>
                        <xs:enumeration value=""ErrorImpersonationDenied""/>
                        <xs:enumeration value=""ErrorImpersonationFailed""/>
                        <xs:enumeration value=""ErrorIncorrectSchemaVersion""/>
                        <xs:enumeration value=""ErrorIncorrectUpdatePropertyCount""/>
                        <xs:enumeration value=""ErrorIndividualMailboxLimitReached""/>
                        <xs:enumeration value=""ErrorInsufficientResources""/>
                        <xs:enumeration value=""ErrorInternalServerError""/>
                        <xs:enumeration value=""ErrorInternalServerTransientError""/>
                        <xs:enumeration value=""ErrorInvalidAccessLevel""/>
                        <xs:enumeration value=""ErrorInvalidArgument""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentId""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentSubfilter""/>
                        <xs:enumeration value=""ErrorInvalidAttachmentSubfilterTextFilter""/>
                        <xs:enumeration value=""ErrorInvalidAuthorizationContext""/>
                        <xs:enumeration value=""ErrorInvalidChangeKey""/>
                        <xs:enumeration value=""ErrorInvalidClientSecurityContext""/>
                        <xs:enumeration value=""ErrorInvalidCompleteDate""/>
                        <xs:enumeration value=""ErrorInvalidContactEmailAddress""/>
                        <xs:enumeration value=""ErrorInvalidContactEmailIndex""/>
                        <xs:enumeration value=""ErrorInvalidCrossForestCredentials""/>
                        <xs:enumeration value=""ErrorInvalidDelegatePermission""/>
                        <xs:enumeration value=""ErrorInvalidDelegateUserId""/>
                        <xs:enumeration value=""ErrorInvalidExcludesRestriction""/>
                        <xs:enumeration value=""ErrorInvalidExpressionTypeForSubFilter""/>
                        <xs:enumeration value=""ErrorInvalidExtendedProperty""/>
                        <xs:enumeration value=""ErrorInvalidExtendedPropertyValue""/>
                        <xs:enumeration value=""ErrorInvalidFolderId""/>
                        <xs:enumeration value=""ErrorInvalidFolderTypeForOperation""/>
                        <xs:enumeration value=""ErrorInvalidFractionalPagingParameters""/>
                        <xs:enumeration value=""ErrorInvalidFreeBusyViewType""/>
                        <xs:enumeration value=""ErrorInvalidId""/>
                        <xs:enumeration value=""ErrorInvalidIdEmpty""/>
                        <xs:enumeration value=""ErrorInvalidIdMalformed""/>
                        <xs:enumeration value=""ErrorInvalidIdMalformedEwsLegacyIdFormat""/>
                        <xs:enumeration value=""ErrorInvalidIdMonikerTooLong""/>
                        <xs:enumeration value=""ErrorInvalidIdNotAnItemAttachmentId""/>
                        <xs:enumeration value=""ErrorInvalidIdReturnedByResolveNames""/>
                        <xs:enumeration value=""ErrorInvalidIdStoreObjectIdTooLong""/>
                        <xs:enumeration value=""ErrorInvalidIdTooManyAttachmentLevels""/>
                        <xs:enumeration value=""ErrorInvalidIdXml""/>
                        <xs:enumeration value=""ErrorInvalidImContactId""/>
                        <xs:enumeration value=""ErrorInvalidImDistributionGroupSmtpAddress""/>
                        <xs:enumeration value=""ErrorInvalidImGroupId""/>
                        <xs:enumeration value=""ErrorInvalidIndexedPagingParameters""/>
                        <xs:enumeration value=""ErrorInvalidInternetHeaderChildNodes""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationArchiveItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCreateItemAttachment""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCreateItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationAcceptItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationDeclineItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationCancelItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationExpandDL""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationRemoveItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationSendItem""/>
                        <xs:enumeration value=""ErrorInvalidItemForOperationTentative""/>
                        <xs:enumeration value=""ErrorInvalidLogonType""/>
                        <xs:enumeration value=""ErrorInvalidMailbox""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderProperty""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderQuota""/>
                        <xs:enumeration value=""ErrorInvalidManagedFolderSize""/>
                        <xs:enumeration value=""ErrorInvalidMergedFreeBusyInterval""/>
                        <xs:enumeration value=""ErrorInvalidNameForNameResolution""/>
                        <xs:enumeration value=""ErrorInvalidOperation""/>
                        <xs:enumeration value=""ErrorInvalidNetworkServiceContext""/>
                        <xs:enumeration value=""ErrorInvalidOofParameter""/>
                        <xs:enumeration value=""ErrorInvalidPagingMaxRows""/>
                        <xs:enumeration value=""ErrorInvalidParentFolder""/>
                        <xs:enumeration value=""ErrorInvalidPercentCompleteValue""/>
                        <xs:enumeration value=""ErrorInvalidPermissionSettings""/>
                        <xs:enumeration value=""ErrorInvalidPhoneCallId""/>
                        <xs:enumeration value=""ErrorInvalidPhoneNumber""/>
                        <xs:enumeration value=""ErrorInvalidUserInfo""/>
                        <xs:enumeration value=""ErrorInvalidPropertyAppend""/>
                        <xs:enumeration value=""ErrorInvalidPropertyDelete""/>
                        <xs:enumeration value=""ErrorInvalidPropertyForExists""/>
                        <xs:enumeration value=""ErrorInvalidPropertyForOperation""/>
                        <xs:enumeration value=""ErrorInvalidPropertyRequest""/>
                        <xs:enumeration value=""ErrorInvalidPropertySet""/>
                        <xs:enumeration value=""ErrorInvalidPropertyUpdateSentMessage""/>
                        <xs:enumeration value=""ErrorInvalidProxySecurityContext""/>
                        <xs:enumeration value=""ErrorInvalidPullSubscriptionId""/>
                        <xs:enumeration value=""ErrorInvalidPushSubscriptionUrl""/>
                        <xs:enumeration value=""ErrorInvalidRecipients""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilter""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterComparison""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterOrder""/>
                        <xs:enumeration value=""ErrorInvalidRecipientSubfilterTextFilter""/>
                        <xs:enumeration value=""ErrorInvalidReferenceItem""/>
                        <xs:enumeration value=""ErrorInvalidRequest""/>
                        <xs:enumeration value=""ErrorInvalidRestriction""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagTypeMismatch""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagInvisible""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagIdGuid""/>
                        <xs:enumeration value=""ErrorInvalidRetentionTagInheritance""/>
                        <xs:enumeration value=""ErrorInvalidRoutingType""/>
                        <xs:enumeration value=""ErrorInvalidScheduledOofDuration""/>
                        <xs:enumeration value=""ErrorInvalidSchemaVersionForMailboxVersion""/>
                        <xs:enumeration value=""ErrorInvalidSecurityDescriptor""/>
                        <xs:enumeration value=""ErrorInvalidSendItemSaveSettings""/>
                        <xs:enumeration value=""ErrorInvalidSerializedAccessToken""/>
                        <xs:enumeration value=""ErrorInvalidServerVersion""/>
                        <xs:enumeration value=""ErrorInvalidSid""/>
                        <xs:enumeration value=""ErrorInvalidSIPUri""/>
                        <xs:enumeration value=""ErrorInvalidSmtpAddress""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterType""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterTypeNotAttendeeType""/>
                        <xs:enumeration value=""ErrorInvalidSubfilterTypeNotRecipientType""/>
                        <xs:enumeration value=""ErrorInvalidSubscription""/>
                        <xs:enumeration value=""ErrorInvalidSubscriptionRequest""/>
                        <xs:enumeration value=""ErrorInvalidSyncStateData""/>
                        <xs:enumeration value=""ErrorInvalidTimeInterval""/>
                        <xs:enumeration value=""ErrorInvalidUserOofSettings""/>
                        <xs:enumeration value=""ErrorInvalidUserPrincipalName""/>
                        <xs:enumeration value=""ErrorInvalidUserSid""/>
                        <xs:enumeration value=""ErrorInvalidUserSidMissingUPN""/>
                        <xs:enumeration value=""ErrorInvalidValueForProperty""/>
                        <xs:enumeration value=""ErrorInvalidWatermark""/>
                        <xs:enumeration value=""ErrorIPGatewayNotFound""/>
                        <xs:enumeration value=""ErrorIrresolvableConflict""/>
                        <xs:enumeration value=""ErrorItemCorrupt""/>
                        <xs:enumeration value=""ErrorItemNotFound""/>
                        <xs:enumeration value=""ErrorItemPropertyRequestFailed""/>
                        <xs:enumeration value=""ErrorItemSave""/>
                        <xs:enumeration value=""ErrorItemSavePropertyError""/>
                        <xs:enumeration value=""ErrorLegacyMailboxFreeBusyViewTypeNotMerged""/>
                        <xs:enumeration value=""ErrorLocalServerObjectNotFound""/>
                        <xs:enumeration value=""ErrorLogonAsNetworkServiceFailed""/>
                        <xs:enumeration value=""ErrorMailboxConfiguration""/>
                        <xs:enumeration value=""ErrorMailboxDataArrayEmpty""/>
                        <xs:enumeration value=""ErrorMailboxDataArrayTooBig""/>
                        <xs:enumeration value=""ErrorMailboxHoldNotFound""/>
                        <xs:enumeration value=""ErrorMailboxLogonFailed""/>
                        <xs:enumeration value=""ErrorMailboxMoveInProgress""/>
                        <xs:enumeration value=""ErrorMailboxStoreUnavailable""/>
                        <xs:enumeration value=""ErrorMailRecipientNotFound""/>
                        <xs:enumeration value=""ErrorMailTipsDisabled""/>
                        <xs:enumeration value=""ErrorManagedFolderAlreadyExists""/>
                        <xs:enumeration value=""ErrorManagedFolderNotFound""/>
                        <xs:enumeration value=""ErrorManagedFoldersRootFailure""/>
                        <xs:enumeration value=""ErrorMeetingSuggestionGenerationFailed""/>
                        <xs:enumeration value=""ErrorMessageDispositionRequired""/>
                        <xs:enumeration value=""ErrorMessageSizeExceeded""/>
                        <xs:enumeration value=""ErrorMimeContentConversionFailed""/>
                        <xs:enumeration value=""ErrorMimeContentInvalid""/>
                        <xs:enumeration value=""ErrorMimeContentInvalidBase64String""/>
                        <xs:enumeration value=""ErrorMissingArgument""/>
                        <xs:enumeration value=""ErrorMissingEmailAddress""/>
                        <xs:enumeration value=""ErrorMissingEmailAddressForManagedFolder""/>
                        <xs:enumeration value=""ErrorMissingInformationEmailAddress""/>
                        <xs:enumeration value=""ErrorMissingInformationReferenceItemId""/>
                        <xs:enumeration value=""ErrorMissingItemForCreateItemAttachment""/>
                        <xs:enumeration value=""ErrorMissingManagedFolderId""/>
                        <xs:enumeration value=""ErrorMissingRecipients""/>
                        <xs:enumeration value=""ErrorMissingUserIdInformation""/>
                        <xs:enumeration value=""ErrorMoreThanOneAccessModeSpecified""/>
                        <xs:enumeration value=""ErrorMoveCopyFailed""/>
                        <xs:enumeration value=""ErrorMoveDistinguishedFolder""/>
                        <xs:enumeration value=""ErrorMultiLegacyMailboxAccess""/>
                        <xs:enumeration value=""ErrorNameResolutionMultipleResults""/>
                        <xs:enumeration value=""ErrorNameResolutionNoMailbox""/>
                        <xs:enumeration value=""ErrorNameResolutionNoResults""/>
                        <xs:enumeration value=""ErrorNoApplicableProxyCASServersAvailable""/>
                        <xs:enumeration value=""ErrorNoCalendar""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToKerberosRequirements""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToSSLRequirements""/>
                        <xs:enumeration value=""ErrorNoDestinationCASDueToVersionMismatch""/>
                        <xs:enumeration value=""ErrorNoFolderClassOverride""/>
                        <xs:enumeration value=""ErrorNoFreeBusyAccess""/>
                        <xs:enumeration value=""ErrorNonExistentMailbox""/>
                        <xs:enumeration value=""ErrorNonPrimarySmtpAddress""/>
                        <xs:enumeration value=""ErrorNoPropertyTagForCustomProperties""/>
                        <xs:enumeration value=""ErrorNoPublicFolderReplicaAvailable""/>
                        <xs:enumeration value=""ErrorNoPublicFolderServerAvailable""/>
                        <xs:enumeration value=""ErrorNoRespondingCASInDestinationSite""/>
                        <xs:enumeration value=""ErrorNotDelegate""/>
                        <xs:enumeration value=""ErrorNotEnoughMemory""/>
                        <xs:enumeration value=""ErrorObjectTypeChanged""/>
                        <xs:enumeration value=""ErrorOccurrenceCrossingBoundary""/>
                        <xs:enumeration value=""ErrorOccurrenceTimeSpanTooBig""/>
                        <xs:enumeration value=""ErrorOperationNotAllowedWithPublicFolderRoot""/>
                        <xs:enumeration value=""ErrorParentFolderIdRequired""/>
                        <xs:enumeration value=""ErrorParentFolderNotFound""/>
                        <xs:enumeration value=""ErrorPasswordChangeRequired""/>
                        <xs:enumeration value=""ErrorPasswordExpired""/>
                        <xs:enumeration value=""ErrorPhoneNumberNotDialable""/>
                        <xs:enumeration value=""ErrorPropertyUpdate""/>
                        <xs:enumeration value=""ErrorPromptPublishingOperationFailed""/>
                        <xs:enumeration value=""ErrorPropertyValidationFailure""/>
                        <xs:enumeration value=""ErrorProxiedSubscriptionCallFailure""/>
                        <xs:enumeration value=""ErrorProxyCallFailed""/>
                        <xs:enumeration value=""ErrorProxyGroupSidLimitExceeded""/>
                        <xs:enumeration value=""ErrorProxyRequestNotAllowed""/>
                        <xs:enumeration value=""ErrorProxyRequestProcessingFailed""/>
                        <xs:enumeration value=""ErrorProxyServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorProxyTokenExpired""/>
                        <xs:enumeration value=""ErrorPublicFolderMailboxDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderOperationFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderRequestProcessingFailed""/>
                        <xs:enumeration value=""ErrorPublicFolderServerNotFound""/>
                        <xs:enumeration value=""ErrorPublicFolderSyncException""/>
                        <xs:enumeration value=""ErrorQueryFilterTooLong""/>
                        <xs:enumeration value=""ErrorQuotaExceeded""/>
                        <xs:enumeration value=""ErrorReadEventsFailed""/>
                        <xs:enumeration value=""ErrorReadReceiptNotPending""/>
                        <xs:enumeration value=""ErrorRecurrenceEndDateTooBig""/>
                        <xs:enumeration value=""ErrorRecurrenceHasNoOccurrence""/>
                        <xs:enumeration value=""ErrorRemoveDelegatesFailed""/>
                        <xs:enumeration value=""ErrorRequestAborted""/>
                        <xs:enumeration value=""ErrorRequestStreamTooBig""/>
                        <xs:enumeration value=""ErrorRequiredPropertyMissing""/>
                        <xs:enumeration value=""ErrorResolveNamesInvalidFolderType""/>
                        <xs:enumeration value=""ErrorResolveNamesOnlyOneContactsFolderAllowed""/>
                        <xs:enumeration value=""ErrorResponseSchemaValidation""/>
                        <xs:enumeration value=""ErrorRestrictionTooLong""/>
                        <xs:enumeration value=""ErrorRestrictionTooComplex""/>
                        <xs:enumeration value=""ErrorResultSetTooBig""/>
                        <xs:enumeration value=""ErrorInvalidExchangeImpersonationHeaderData""/>
                        <xs:enumeration value=""ErrorSavedItemFolderNotFound""/>
                        <xs:enumeration value=""ErrorSchemaValidation""/>
                        <xs:enumeration value=""ErrorSearchFolderNotInitialized""/>
                        <xs:enumeration value=""ErrorSendAsDenied""/>
                        <xs:enumeration value=""ErrorSendMeetingCancellationsRequired""/>
                        <xs:enumeration 
                            value=""ErrorSendMeetingInvitationsOrCancellationsRequired""/>
                        <xs:enumeration value=""ErrorSendMeetingInvitationsRequired""/>
                        <xs:enumeration value=""ErrorSentMeetingRequestUpdate""/>
                        <xs:enumeration value=""ErrorSentTaskRequestUpdate""/>
                        <xs:enumeration value=""ErrorServerBusy""/>
                        <xs:enumeration value=""ErrorServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorStaleObject""/>
                        <xs:enumeration value=""ErrorSubmissionQuotaExceeded""/>
                        <xs:enumeration value=""ErrorSubscriptionAccessDenied""/>
                        <xs:enumeration value=""ErrorSubscriptionDelegateAccessNotSupported""/>
                        <xs:enumeration value=""ErrorSubscriptionNotFound""/>
                        <xs:enumeration value=""ErrorSubscriptionUnsubscribed""/>
                        <xs:enumeration value=""ErrorSyncFolderNotFound""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotFound""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotLinkedToSharePoint""/>
                        <xs:enumeration value=""ErrorTeamMailboxUrlValidationFailed""/>
                        <xs:enumeration value=""ErrorTeamMailboxNotAuthorizedOwner""/>
                        <xs:enumeration value=""ErrorTeamMailboxActiveToPendingDelete""/>
                        <xs:enumeration value=""ErrorTeamMailboxFailedSendingNotifications""/>
                        <xs:enumeration value=""ErrorTeamMailboxErrorUnknown""/>
                        <xs:enumeration value=""ErrorTimeIntervalTooBig""/>
                        <xs:enumeration value=""ErrorTimeoutExpired""/>
                        <xs:enumeration value=""ErrorTimeZone""/>
                        <xs:enumeration value=""ErrorToFolderNotFound""/>
                        <xs:enumeration value=""ErrorTokenSerializationDenied""/>
                        <xs:enumeration value=""ErrorUpdatePropertyMismatch""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingDialPlanNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingReportDataNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingPromptNotFound""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingRequestFailed""/>
                        <xs:enumeration value=""ErrorUnifiedMessagingServerNotFound""/>
                        <xs:enumeration value=""ErrorUnableToGetUserOofSettings""/>
                        <xs:enumeration value=""ErrorUnableToRemoveImContactFromGroup""/>
                        <xs:enumeration value=""ErrorUnsupportedSubFilter""/>
                        <xs:enumeration value=""ErrorUnsupportedCulture""/>
                        <xs:enumeration value=""ErrorUnsupportedMapiPropertyType""/>
                        <xs:enumeration value=""ErrorUnsupportedMimeConversion""/>
                        <xs:enumeration value=""ErrorUnsupportedPathForQuery""/>
                        <xs:enumeration value=""ErrorUnsupportedPathForSortGroup""/>
                        <xs:enumeration value=""ErrorUnsupportedPropertyDefinition""/>
                        <xs:enumeration value=""ErrorUnsupportedQueryFilter""/>
                        <xs:enumeration value=""ErrorUnsupportedRecurrence""/>
                        <xs:enumeration value=""ErrorUnsupportedTypeForConversion""/>
                        <xs:enumeration value=""ErrorUpdateDelegatesFailed""/>
                        <xs:enumeration value=""ErrorUserNotUnifiedMessagingEnabled""/>
                        <xs:enumeration value=""ErrorValueOutOfRange""/>
                        <xs:enumeration value=""ErrorVoiceMailNotImplemented""/>
                        <xs:enumeration value=""ErrorVirusDetected""/>
                        <xs:enumeration value=""ErrorVirusMessageDeleted""/>
                        <xs:enumeration value=""ErrorWebRequestInInvalidState""/>
                        <xs:enumeration value=""ErrorWin32InteropError""/>
                        <xs:enumeration value=""ErrorWorkingHoursSaveFailed""/>
                        <xs:enumeration value=""ErrorWorkingHoursXmlMalformed""/>
                        <xs:enumeration value=""ErrorWrongServerVersion""/>
                        <xs:enumeration value=""ErrorWrongServerVersionDelegate""/>
                        <xs:enumeration value=""ErrorMissingInformationSharingFolderId""/>
                        <xs:enumeration value=""ErrorDuplicateSOAPHeader""/>
                        <xs:enumeration value=""ErrorSharingSynchronizationFailed""/>
                        <xs:enumeration value=""ErrorSharingNoExternalEwsAvailable""/>
                        <xs:enumeration value=""ErrorFreeBusyDLLimitReached""/>
                        <xs:enumeration value=""ErrorInvalidGetSharingFolderRequest""/>
                        <xs:enumeration value=""ErrorNotAllowedExternalSharingByPolicy""/>
                        <xs:enumeration value=""ErrorUserNotAllowedByPolicy""/>
                        <xs:enumeration value=""ErrorPermissionNotAllowedByPolicy""/>
                        <xs:enumeration value=""ErrorOrganizationNotFederated""/>
                        <xs:enumeration value=""ErrorMailboxFailover""/>
                        <xs:enumeration value=""ErrorInvalidExternalSharingInitiator""/>
                        <xs:enumeration value=""ErrorMessageTrackingPermanentError""/>
                        <xs:enumeration value=""ErrorMessageTrackingTransientError""/>
                        <xs:enumeration value=""ErrorMessageTrackingNoSuchDomain""/>
                        <xs:enumeration value=""ErrorUserWithoutFederatedProxyAddress""/>
                        <xs:enumeration value=""ErrorInvalidOrganizationRelationshipForFreeBusy""/>
                        <xs:enumeration value=""ErrorInvalidFederatedOrganizationId""/>
                        <xs:enumeration value=""ErrorInvalidExternalSharingSubscriber""/>
                        <xs:enumeration value=""ErrorInvalidSharingData""/>
                        <xs:enumeration value=""ErrorInvalidSharingMessage""/>
                        <xs:enumeration value=""ErrorNotSupportedSharingMessage""/>
                        <xs:enumeration value=""ErrorApplyConversationActionFailed""/>
                        <xs:enumeration value=""ErrorInboxRulesValidationError""/>
                        <xs:enumeration value=""ErrorOutlookRuleBlobExists""/>
                        <xs:enumeration value=""ErrorRulesOverQuota""/>
                        <xs:enumeration value=""ErrorNewEventStreamConnectionOpened""/>
                        <xs:enumeration value=""ErrorMissedNotificationEvents""/>
                        <xs:enumeration value=""ErrorDuplicateLegacyDistinguishedName""/>
                        <xs:enumeration value=""ErrorInvalidClientAccessTokenRequest""/>
                        <xs:enumeration value=""ErrorNoSpeechDetected""/>
                        <xs:enumeration value=""ErrorUMServerUnavailable""/>
                        <xs:enumeration value=""ErrorRecipientNotFound""/>
                        <xs:enumeration value=""ErrorRecognizerNotInstalled""/>
                        <xs:enumeration value=""ErrorSpeechGrammarError""/>
                        <xs:enumeration value=""ErrorInvalidManagementRoleHeader""/>
                        <xs:enumeration value=""ErrorLocationServicesDisabled""/>
                        <xs:enumeration value=""ErrorLocationServicesRequestTimedOut""/>
                        <xs:enumeration value=""ErrorLocationServicesRequestFailed""/>
                        <xs:enumeration value=""ErrorLocationServicesInvalidRequest""/>
                        <xs:enumeration value=""ErrorMailboxScopeNotAllowedWithoutQueryString""/>
                        <xs:enumeration value=""ErrorArchiveMailboxSearchFailed""/>
                        <xs:enumeration value=""ErrorArchiveMailboxServiceDiscoveryFailed""/>
                        <xs:enumeration value=""ErrorInvalidPhotoSize""/>
                        <xs:enumeration value=""ErrorSearchQueryHasTooManyKeywords""/>
                        <xs:enumeration value=""ErrorSearchTooManyMailboxes""/>
                        <xs:enumeration value=""ErrorDiscoverySearchesDisabled""/>
                    </xs:restriction>
                </xs:simpleType>");
            }

            if (this.exchangeServiceBinding.ServerVersionInfoValue != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1339");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1339
                Site.CaptureRequirementIfIsTrue(
                    isSchemaValidated,
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
            }
        }
    }
}