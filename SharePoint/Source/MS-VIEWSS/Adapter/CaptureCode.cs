namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System.Web.Services.Protocols;
    using System.Xml;
    using Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_VIEWSSAdapter : ManagedAdapterBase, IMS_VIEWSSAdapter
    {
        /// <summary>
        /// Gets a value indicating whether the schema validation of the server response is valid.
        /// </summary>
        private bool PassSchemaValidation
        {
            get
            {
                if ((SchemaValidation.ValidationResult != ValidationResult.Error)
                    && (SchemaValidation.ValidationResult != ValidationResult.Warning))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// This method is used to capture requirements of SOAPFaultDetails complex type.
        /// </summary>
        /// <param name="detail">The detail element in a SOAP Fault.</param>
        private void ValidateSOAPFaultDetails(XmlNode detail)
        {
            string soapFaultDetailsElement = SchemaValidation.GetSoapFaultDetailBody(detail.OuterXml);
            Site.Log.Add(LogEntryKind.Debug, "The SOAP fault details element is: \r\n{0}", soapFaultDetailsElement);
            ValidationResult detailsValidationResult = SchemaValidation.ValidateXml(this.Site, soapFaultDetailsElement);
            bool isDetailValidated = false;
            if ((detailsValidationResult != ValidationResult.Error)
                && (detailsValidationResult != ValidationResult.Warning))
            {
                isDetailValidated = true;
            }

            // If the schema validation of the SOAP fault detail succeeds, then MS-VIEWSS_R5 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isDetailValidated,
                5,
                @"[In Transport] Protocol server faults MUST be returned [either via HTTP status codes, as specified in [RFC2616] section 10 (Status Code Definitions), or] via SOAP faults, as specified either in [SOAP1.1] section 4.4 (SOAP Fault) or in [SOAP1.2/1] section 5.4 (SOAP Fault).");

            // If the schema validation succeeds, then R159 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isDetailValidated,
                159,
                @"[In SOAPFaultDetails] The definition of the SOAPFaultDetails element is as follows.
                                        <s:schema xmlns:s=""http://www.w3.org/2001/XMLSchema"" targetNamespace="" http://schemas.microsoft.com/sharepoint/soap/"">
                                           <s:complexType name=""SOAPFaultDetails"">
                                              <s:sequence>
                                                 <s:element name=""errorstring"" type=""s:string""/>
                                                 <s:element name=""errorcode"" type=""s:string"" minOccurs=""0""/>
                                              </s:sequence>
                                           </s:complexType>
                                        </s:schema>");

            if (this.viewssProxy.SoapVersion == SoapProtocolVersion.Soap11)
            {
                // If the soap version is 1.1, the soap detail element exists and the schema validation succeeds, then R8901 can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isDetailValidated,
                    8901,
                    "[In Protocol Details] This protocol allows protocol servers to provide additional details for SOAP faults by including a detail element as specified in [SOAP1.1] section 4.4, which conforms to the XML schema of the SOAPFaultDetails complex type specified in section 2.2.4.2.");
            }
            else if (this.viewssProxy.SoapVersion == SoapProtocolVersion.Soap12)
            {
                // If the soap version is 1.2, the soap detail element exists and the schema validation succeeds, then R8902001 can be captured.
                Site.CaptureRequirementIfIsTrue(
                    isDetailValidated,
                    8902001,
                    "[In Appendix B: Product Behavior] Implementation does use a detail element instead of the Detail element in SOAP 1.2. <1> Section 3:  Microsoft products use a detail element instead of the Detail element in SOAP 1.2.");
            }

            XmlNodeList childNodes = detail.ChildNodes;
            this.Site.Assert.IsTrue(childNodes.Count == 1 || childNodes.Count == 2, "This XmlNode must have one or two child nodes, and errorcode node's minOccurs = 0.");

            XmlNode node1 = childNodes.Item(0);
            this.Site.Assert.AreEqual("errorstring", node1.Name, "The element errorstring should exist as the first child in the detail element.");
            this.Site.Assert.IsTrue(!string.IsNullOrEmpty(node1.CreateNavigator().Value), "The element errorstring should not be empty to contain a human readable text explaining the application-level fault.");

            if (childNodes.Count == 2)
            {
                XmlNode node2 = childNodes.Item(1);
                this.Site.Assert.AreEqual("errorcode", node2.Name, "The second child node should be errorcode.");

                string errorCode = node2.CreateNavigator().Value;
                bool isErrorCodeHexadecimal = this.IsErrorCodeHexadecimal(errorCode);
                Site.CaptureRequirementIfIsTrue(
                    isErrorCodeHexadecimal,
                    41,
                    "[In SOAPFaultDetails] errorcode: The hexadecimal representation of a 4-byte result code.");
            }
        }

        /// <summary>
        /// Capture underlying transport protocol related requirements.
        /// </summary>
        private void CaptureTransportRelatedRequirements()
        {
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            switch (transport)
            {
                case TransportProtocol.HTTP:

                    // As response successfully returned, the transport related requirements can be captured.
                    Site.CaptureRequirement(
                        1,
                        @"[In Transport]Protocol servers MUST support SOAP over HTTP.");
                    break;

                case TransportProtocol.HTTPS:

                    if (Common.IsRequirementEnabled(433, this.Site))
                    {
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                            433,
                            @"[In Appendix B: Product Behavior] Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients.(Windows SharePoint Services 2.0 and above products follow this behavior.)");
                    }

                    break;

                default:
                    Site.Debug.Fail("Unknown transport type " + transport);
                    break;
            }

            SoapProtocolVersion soapVersion = this.viewssProxy.SoapVersion;

            // Verifies MS-VIEWSS requirement: MS-VIEWSS_R4.
            Site.CaptureRequirementIfIsTrue(
                soapVersion == SoapProtocolVersion.Soap11 || soapVersion == SoapProtocolVersion.Soap12,
                4,
                @"[In Transport] Protocol messages MUST be formatted as specified either in [SOAP1.1] section 4 (SOAP Envelope) or in [SOAP1.2/1] section 5 (SOAP Message Construct).");
        }

        /// <summary>
        /// Used to validate the available query information related requirements.
        /// </summary>
        /// <param name="queryRoot">The returned view query instance.</param>
        private void ValidateQuery(CamlQueryRoot queryRoot)
        {
            // If the query is not null, its schema has been validated before de-serialization; so the requirement MS-WSSCAML_R20 can be captured.
            if (queryRoot != null)
            {
                Site.CaptureRequirementIfIsTrue(
                    this.PassSchemaValidation,
                    "MS-WSSCAML",
                    20,
                    @"[In Schema] [The schema definition of CamlQueryRoot Type is as follows:]
                                                        <xs:complexType name=""CamlQueryRoot"">
                                                        <xs:all>
                                                        <xs:element name=""Where"" type=""LogicalJoinDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                        <xs:element name=""OrderBy"" type=""OrderByDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                        <xs:element name=""GroupBy"" type=""GroupByDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                        <xs:element name=""WithIndex"" type=""LogicalWithIndexDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                        </xs:all>
                                                        </xs:complexType>");

                if (queryRoot.OrderBy != null)
                {
                    // If ViewFields contains any element, its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R56 can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        56,
                        @"[In Schema] [The schema definition of OrderByDefinition is as follows]  <xs:complexType name=""OrderByDefinition"">
                                                          <xs:sequence>
                                                          <xs:element name=""FieldRef"" type=""FieldRefDefinitionOrderBy"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                                          </xs:sequence>
                                                          <xs:attribute name=""Override"" type=""TRUE_Case_Insensitive_Else_Anything"" use=""optional"" default=""FALSE""/>
                                                          </xs:complexType>");

                    if (queryRoot.OrderBy.FieldRef != null && queryRoot.OrderBy.FieldRef.Length > 0)
                    {
                        // If ViewFields contains any element, and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R59 can be captured.
                        Site.CaptureRequirementIfIsTrue(
                            this.PassSchemaValidation,
                            "MS-WSSCAML",
                            59,
                            @"[In Schema] [The schema definition of FieldRefDefinitionOrderBy is as follows:]<xs:complexType name=""FieldRefDefinitionOrderBy"">
                                                              <xs:attribute name=""ID"" type=""UniqueIdentifierWithOrWithoutBraces"" use=""optional"" />
                                                              <xs:attribute name=""Name"" type=""xs:string"" use=""optional"" />
                                                              <xs:attribute name=""Ascending"" type=""TRUE_Case_Insensitive_Else_Anything"" use=""optional"" default=""FALSE"" />
                                                              </xs:complexType>");

                        foreach (FieldRefDefinitionOrderBy fieldRef in queryRoot.OrderBy.FieldRef)
                        {
                            if (fieldRef.Ascending != null)
                            {
                                // If Ascending exists, and its schema has been validated before DE serializing So this requirement MS-WSSCAML_R1455 can be captured.
                                Site.CaptureRequirementIfIsTrue(
                                    this.PassSchemaValidation,
                                    "MS-WSSCAML",
                                    1455,
                                    @"[In TRUE_Case_Insensitive_Else_Anything] This type is defined as follows:
                                                                              <xs:simpleType name=""TRUE_Case_Insensitive_Else_Anything"">
                                                                                <xs:restriction base=""xs:string"">
                                                                                  <xs:pattern value=""([Tt][Rr][Uu][Ee])|.*"" />
                                                                                </xs:restriction>
                                                                              </xs:simpleType>");
                            }
                        }
                    }
                }

                // If the Where element is not null, its schema type LogicalJoinDefinition has been validated before de-serialization; so this requirement MS-WSSCAML_R24 can be captured.
                if (queryRoot.Where != null)
                {
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        24,
                        @"[In Schema] [The schema definition of LogicalJoinDefinition type is as follows:] <xs:complexType name=""LogicalJoinDefinition"">
                                                                                              <xs:choice minOccurs=""0"" maxOccurs=""unbounded"">
                                                                                              <xs:element name=""And"" type=""ExtendedLogicalJoinDefinition"" />
                                                                                              <xs:element name=""BeginsWith"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Contains"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""DateRangesOverlap"" type=""LogicalTestDefinitionDateRange"" />
                                                                                              <xs:element name=""Eq"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Geq"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Gt"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""In"" type=""LogicalTestInValuesDefinition"" />
                                                                                              <xs:element name=""Includes"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""IsNotNull"" type=""LogicalNullDefinition"" />
                                                                                              <xs:element name=""IsNull"" type=""LogicalNullDefinition"" />
                                                                                              <xs:element name=""Leq"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Lt"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Membership"" type=""MembershipDefinition"" />
                                                                                              <xs:element name=""Neq"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""NotIncludes"" type=""LogicalTestDefinition"" />
                                                                                              <xs:element name=""Or"" type=""ExtendedLogicalJoinDefinition"" />
                                                                                              </xs:choice>
                                                                                              </xs:complexType>");
                    switch (queryRoot.Where.ItemsElementName[0])
                    {
                        case ItemsChoiceType1.Contains:
                        case ItemsChoiceType1.BeginsWith:
                        case ItemsChoiceType1.Eq:
                        case ItemsChoiceType1.Geq:
                        case ItemsChoiceType1.Gt:
                        case ItemsChoiceType1.Includes:
                        case ItemsChoiceType1.NotIncludes:
                        case ItemsChoiceType1.Leq:
                        case ItemsChoiceType1.Lt:
                        case ItemsChoiceType1.Neq:

                            // If the Where element exists and used one of the above comparing signs that is defined by the LogicalTestDefinition complex type and the schema of the comparing sign has been validated before de-serialization, the requirement MS-WSSCAML_R79 can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                this.PassSchemaValidation,
                                "MS-WSSCAML",
                                79,
                                @"[In Schema] [The schema definition of LogicalTestDefinition is as follows:] 
                                             <xs:complexType name=""LogicalTestDefinition"">
                                                <xs:all>
                                                  <xs:element name=""FieldRef"" type=""FieldRefDefinitionQueryTest"" minOccurs=""1"" maxOccurs=""1"" />
                                                  <xs:element name=""Value"" type=""ValueDefinition"" minOccurs=""1"" maxOccurs=""1"" />
                                                </xs:all>
                                              </xs:complexType>");

                            // Since the Value element is required to exist by its schema definition (see the above MS-WSSCAML_R79),  and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R82 can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                 this.PassSchemaValidation,
                                 "MS-WSSCAML",
                                 82,
                                 @"[In Schema] [The schema definition of ValueDefinition is as follows:]
                                                        <xs:complexType name=""ValueDefinition"" mixed=""true"">
                                                        <xs:sequence>
                                                        <xs:choice minOccurs=""0"" maxOccurs=""unbounded"">
                                                        <xs:any namespace=""##any"" processContents=""skip"" />
                                                        </xs:choice>
                                                        </xs:sequence>
                                                        <xs:attribute name=""Type"" type=""xs:string"" use=""optional"" />
                                                        </xs:complexType>");

                            // Since FieldRef is required to exist by its schema definition (see the above MS-WSSCAML_R79), and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R87 can be captured.
                            Site.CaptureRequirementIfIsTrue(
                                 this.PassSchemaValidation,
                                 "MS-WSSCAML",
                                 87,
                                 @"[In Schema] [The schema definition of FieldRefDefinitionQueryTest is as follows:] 
                                                        <xs:complexType name=""FieldRefDefinitionQueryTest"">
                                                        <xs:attribute name=""ID"" type="" UniqueIdentifierWithOrWithoutBraces"" use=""optional"" />
                                                        <xs:attribute name=""Name"" type=""xs:string"" use=""optional"" />
                                                        <xs:attribute name=""LookupId"" type=""TRUE_Case_Insensitive_Else_Anything"" use=""optional"" default=""FALSE"" />
                                                        </xs:complexType>");
                            break;
                        case ItemsChoiceType1.And:
                        case ItemsChoiceType1.Or:
                        case ItemsChoiceType1.In:
                        case ItemsChoiceType1.IsNull:
                        case ItemsChoiceType1.IsNotNull:
                        case ItemsChoiceType1.Membership:
                        case ItemsChoiceType1.DateRangesOverlap:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Used to validate the View Definition
        /// </summary>
        /// <param name="view">specify the type of the view</param>
        private void ValidateViewDefinition(ViewDefinition view)
        {
            if (view != null)
            {
                // If the element view is not null, and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R799 can be captured.           
                Site.CaptureRequirementIfIsTrue(
                    this.PassSchemaValidation,
                    "MS-WSSCAML",
                    799,
                    @"[In Schema] [The schema definition of ViewDefinition type is as follows:]
                                    <xs:complexType name=""ViewDefinition"">
                                        <xs:group ref=""ViewDefinitionChildElementGroup""/> 
                                        <xs:attribute name=""AggregateView"" type=""TRUEFALSE""  default=""FALSE""/>
                                        <xs:attribute name=""BaseViewID"" type=""xs:int"" />
                                        <xs:attribute name=""CssStyleSheet"" type=""xs:string"" />
                                        <xs:attribute name=""DefaultView"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""DisplayName"" type=""xs:string"" />
                                        <xs:attribute name=""FailIfEmpty"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""FileDialog"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""FPModified"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""Hidden"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""List"" type=""UniqueIdentifierWithoutBraces"" />
                                        <xs:attribute name=""Name"" type=""UniqueIdentifierWithBraces"" />
                                        <xs:attribute name=""ContentTypeID"" type=""ContentTypeId"" />
                                        <xs:attribute name=""OrderedView"" type=""TRUEFALSE"" />
                                        <xs:attribute name=""DefaultViewForContentType"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""IncludeRootFolder"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""PageType"" type=""xs:string"" />
                                        <xs:attribute name=""Path"" type=""RelativeFilePath"" />
                                        <xs:attribute name=""Personal"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""ReadOnly"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""RecurrenceRowset"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""RequiresClientIntegration"" type=""TRUEFALSE"" default=""FALSE""/>
                                        <xs:attribute name=""RowLimit"" type=""xs:int"" />
                                        <xs:attribute name=""ShowHeaderUI"" type=""TRUEFALSE""  default=""FALSE"" />
                                        <xs:attribute name=""Type"" type=""ViewType"" default=""HTML""/>
                                        <xs:attribute name=""Url"" type=""RelativeUrl"" />
                                        <xs:attribute name=""UseSchemaXmlToolbar"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""WebPartOrder"" type=""xs:int"" />
                                        <xs:attribute name=""WebPartZoneID"" type=""xs:string"" />
                                        <xs:attribute name=""FreeForm"" type=""TRUEFALSE"" />
                                        <xs:attribute name=""ImageUrl"" type=""xs:string"" />
                                        <xs:attribute name=""SetupPath"" type=""RelativeFilePath"" />
                                        <xs:attribute name=""ToolbarTemplate"" type=""xs:string"" />
                                        <xs:attribute name=""MobileView"" type=""TRUEFALSE"" default=""FALSE""/>
                                        <xs:attribute name=""MobileDefaultView"" type=""TRUEFALSE"" />
                                        <xs:attribute name=""MobileUrl"" type=""RelativeUrl"" />
                                        <xs:attribute name=""Level"" type=""ViewPageLevel"" default=""1"" />
                                        <xs:attribute name=""FrameState"" type=""xs:string"" default=""Normal"" />
                                        <xs:attribute name=""IsIncluded"" type=""TRUEFALSE"" default=""TRUE"" />
                                        <xs:attribute name=""IncludeVersions"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""HackLockWeb"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""ModerationType"" type=""ViewModerationType"" default="""" />
                                        <xs:attribute name=""Scope"" type=""ViewScope"" default="""" />
                                        <xs:attribute name=""Threaded"" type=""TRUEFALSE"" default=""FALSE"" />
                                        <xs:attribute name=""TabularView"" type=""FALSE_Case_Insensitive_Else_Anything"" default=""TRUE"" />
                                      </xs:complexType>

                                      <xs:group name=""ViewDefinitionChildElementGroup"">
                                        <xs:all>
                                          <xs:element name=""PagedRowset"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""Toolbar"" type=""ToolbarDefinition""  minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""Query"" type=""CamlQueryRoot"" minOccurs=""0"" maxOccurs=""1"" />      
                                          <xs:element name=""ViewFields"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:element name=""FieldRef"" type=""FieldRefDefinitionView"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                              </xs:sequence>
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""GroupByHeader"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""GroupByFooter"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""ViewHeader"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""ViewBody"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""ViewFooter"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""RowLimitExceeded"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""ViewEmpty"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""PagedRecurrenceRowset"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""PagedClientCallbackRowset"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""Aggregations"" type=""AggregationsDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""OpenApplicationExtension"" type=""xs:string"" minOccurs=""0"" maxOccurs=""1"" /> 
                                          <xs:element name=""RowLimit"" type=""RowLimitDefinition"" minOccurs=""0"" maxOccurs=""1"" default=""2147483647"" />
                                          <xs:element name=""Mobile"" type=""MobileViewDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""ViewStyle"" type=""ViewStyleReference"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""CalendarSettings"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""CalendarViewStyles"" type=""CalendarViewStylesDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""ViewBidiHeader"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""Script"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""ViewData"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>           
                                                <xs:element name=""FieldRef"" type=""FieldRefDefinitionViewData"" minOccurs=""3"" maxOccurs=""5"" />
                                              </xs:sequence>
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""Formats"" type=""ViewFormatDefinitions"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""InlineEdit"" type=""TRUE_If_Present"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""ProjectedFields"" type=""ProjectedFieldsDefinitionType"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""Joins"" type=""ListJoinsDefinitionType"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""Method"" type=""ViewMethodDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                          <xs:element name=""ParameterBindings"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""Xsl"" minOccurs=""0"" maxOccurs=""1"">
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                          </xs:element>
                                          <xs:element name=""XslLink"" minOccurs=""0"" maxOccurs=""1"" >
                                            <xs:complexType>
                                              <xs:sequence>
                                                <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                              </xs:sequence>
                                              <xs:anyAttribute processContents=""skip"" />
                                            </xs:complexType>
                                            </xs:element>
                                       </xs:all>
                                      </xs:group>");

                if (view.Query != null)
                {
                    // Validate common query requirements.
                    this.ValidateQuery(view.Query);
                }

                if (view.RowLimit != null)
                {
                    // Validate row limit related schema requirements.
                    this.ValidateRowLimit();
                }

                if (view.ViewFields != null && view.ViewFields.Length > 0)
                {
                    // Validate view fields related schema requirements.
                    this.ValidateViewFields();
                }
                
                // If the string of Type equals one of enumeration values of the ViewType type, and its schema has been verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R194 can be captured.
                bool isViewType = view.Type.ToString().Equals("HTML") ||
                     view.Type.ToString().Equals("GRID") ||
                     view.Type.ToString().Equals("CALENDAR") ||
                     view.Type.ToString().Equals("RECURRENCE") ||
                     view.Type.ToString().Equals("CHART") ||
                     view.Type.ToString().Equals("GANTT") ||
                     view.Type.ToString().Equals("TABLE");

                Site.CaptureRequirementIfIsTrue(
                    isViewType,
                    "MS-WSSCAML",
                    194,
                    @"[In ViewType] The ViewType type specifies the type of view rendering.
                                   <xs:simpleType name=""ViewType"">
                                    <xs:restriction base=""xs:string"">
                                      <xs:enumeration value=""HTML"" />
                                      <xs:enumeration value=""GRID"" />
                                      <xs:enumeration value=""CALENDAR"" />
                                      <xs:enumeration value=""RECURRENCE"" />
                                      <xs:enumeration value=""CHART"" />
                                      <xs:enumeration value=""GANTT"" />
                                      <xs:enumeration value=""TABLE"" />      
                                    </xs:restriction>
                                  </xs:simpleType>");

                // If level equals one of the enumeration integer values of the ViewPageLevel type, and its schema has been verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R188 can be captured.
                bool isViewPageLevel = view.Level == 1 ||
                    view.Level == 2 ||
                    view.Level == 255;

                Site.CaptureRequirementIfIsTrue(
                    isViewPageLevel,
                    "MS-WSSCAML",
                    188,
                    @"[In ViewPageLevel] [The schema definition of ViewPageLevel is as follows:]  <xs:simpleType name=""ViewPageLevel"">
                                                                                      <xs:restriction base=""xs:int "">
                                                                                      <xs:enumeration value=""1"" />
                                                                                      <xs:enumeration value=""2"" />
                                                                                      <xs:enumeration value=""255"" />
                                                                                      </xs:restriction>
                                                                                      </xs:simpleType>");

                // If this attribute ContentTypeID exists, and its schema has been verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R140 can be captured.
                if (view.ContentTypeID != null)
                {
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        140,
                        @"[In ContentTypeId Type] The ContentTypeId type is the identifier for the specified content type. See [MS-WSSTS] section 2.1.2.8.1 for more information about the structure of a content type identifier.
                                                                                                <xs:simpleType name=""ContentTypeId"">
                                                                                                <xs:restriction base=""xs:string"">
                                                                                                <xs:pattern value=""0x([0-9A-Fa-f][1-9A-Fa-f]|[1-9A-Fa-f][0-9A-Fa-f]|00[0-9A-Fa-f]{32})*"" />
                                                                                                <xs:minLength value=""2""/>
                                                                                                <xs:maxLength value=""1026""/>
                                                                                                </xs:restriction>
                                                                                                </xs:simpleType>");
                }
                
                // If the attribute Url exists, and its schema has been verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R7 can be captured.
                if (view.Url != null)
                {
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        7,
                        @"[In RelativeUrl] [The schema definition of RelativeUrl is as follows:] 
                                                                                  <xs:simpleType name=""RelativeUrl"">
                                                                                  <xs:restriction base=""xs:string"" >
                                                                                  <xs:maxLength value=""255"" />
                                                                                  <xs:minLength value=""0"" />
                                                                                  </xs:restriction>
                                                                                  </xs:simpleType>");
                }

                // If the attribute Name exists, and its schema has been verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R17 can be captured.
                if (view.Name != null)
                {
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        17,
                        @"[In UniqueIdentifierWithBraces] [The schema definition of UniqueIdentifierWithBraces is as follows:] 
                                                                                          <xs:simpleType name=""UniqueIdentifierWithBraces"">
                                                                                          <xs:restriction base=""xs:string"">
                                                                                          <xs:pattern value=""\{[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}\}""/>
                                                                                          </xs:restriction>
                                                                                          </xs:simpleType>");
                }

                if (view.Toolbar != null)
                {
                    // If Toolbar exists, and its schema is verified as conforming to the schema in WSDL, then this requirement MS-WSSCAML_R788 can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        788,
                        @"[In Schema] [The schema definition of ToolbarDefinition is as follows:]  
                                                                   <xs:complexType name=""ToolbarDefinition"">
                                                                   <xs:sequence>
                                                                   <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any"" processContents=""skip"" />
                                                                   </xs:sequence>
                                                                   <xs:attribute name=""Position"" type=""ToolbarPosition"" />
                                                                   <xs:attribute name=""Type"" type=""ToolbarType"" />
                                                                   <xs:anyAttribute processContents=""skip"" />
                                                                   </xs:complexType>");

                    // If the attribute Type exists, and its schema is verified to conform to the schema in WSDL, then this requirement MS-WSSCAML_R177 can be captured.
                    if (view.Toolbar.TypeSpecified)
                    {
                        Site.CaptureRequirementIfIsTrue(
                            this.PassSchemaValidation,
                            "MS-WSSCAML",
                            177,
                            @"[In ToolbarType] [The schema definition of ToolbarType is as follows:]
                                                 <xs:simpleType name=""ToolbarType"">
                                                    <xs:restriction base=""xs:string"">
                                                      <xs:enumeration value=""Standard"" />
                                                      <xs:enumeration value=""FreeForm"" />
                                                      <xs:enumeration value=""RelatedTasks"" />      
                                                      <xs:enumeration value=""None"" />
                                                    </xs:restriction>
                                                  </xs:simpleType>");
                    }
                }

                this.ValidateFormats(view.Formats);
            }
        }

        /// <summary>
        /// Used to validate the Formats
        /// </summary>
        /// <param name="formats">specify the formats of the views</param>
        private void ValidateFormats(ViewFormatDefinitions formats)
        {
            if (formats != null)
            {
                // If the element formats is not null, and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R900 can be captured.
                Site.CaptureRequirementIfIsTrue(
                    this.PassSchemaValidation,
                    "MS-WSSCAML",
                    900,
                    @"[In Schema] [The schema definition of ViewFormatDefinitions Type is as follows:]
                                                                   <xs:complexType name=""ViewFormatDefinitions"">
                                                                   <xs:sequence>
                                                                   <xs:element name=""FormatDef"" type=""FormatDefDefinition"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                                                   <xs:element name=""Format"" type=""FormatDefinition""  minOccurs=""0"" maxOccurs=""unbounded"" />
                                                                   </xs:sequence>
                                                                   </xs:complexType>");

                if (formats.Format != null)
                {
                    // If the element Format is not null, and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R907 can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        907,
                        @"[In Schema] [The schema definition of FormatDefinition is as follows:]
                                                                      <xs:complexType name=""FormatDefinition"">
                                                                      <xs:sequence>
                                                                      <xs:element name=""FormatDef"" type=""FormatDefDefinition"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                                                      </xs:sequence>
                                                                      <xs:attribute name=""Name"" type=""xs:string"" use=""required"" />
                                                                      </xs:complexType>");
                }

                if (formats.FormatDef != null)
                {
                    // If the element FormatDef is not null, its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R904 can be captured.
                    Site.CaptureRequirementIfIsTrue(
                        this.PassSchemaValidation,
                        "MS-WSSCAML",
                        904,
                        @"[In Schema] [The schema definition of FormatDefDefinition is as follows:] <xs:complexType name=""FormatDefDefinition"">
                                                                  <xs:simpleContent>
                                                                  <xs:extension base=""xs:string"">
                                                                  <xs:attribute name=""Type"" type=""xs:string"" use=""required"" />
                                                                  <xs:attribute name=""Value"" type=""xs:string"" use=""required"" />
                                                                  </xs:extension>
                                                                  </xs:simpleContent>
                                                                  </xs:complexType>");
                }
            }
        }

        /// <summary>
        /// Validate View's RowLimit element related requirements.
        /// </summary>
        private void ValidateRowLimit()
        {
            // If the element RowLimit is not null, its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R784 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                "MS-WSSCAML",
                 784,
                 @"[In Schema] [The schema definition of RowLimitDefinition type is as follows:] 
                                             <xs:complexType name=""RowLimitDefinition"">
                                             <xs:simpleContent>
                                             <xs:extension base=""xs:int"">
                                             <xs:attribute name=""Paged"" type=""TRUE_Case_Insensitive_Else_Anything"" use=""optional"" default=""FALSE""/>
                                             </xs:extension>
                                             </xs:simpleContent>
                                             </xs:complexType>");
        }

        /// <summary>
        /// Validate View's Fields element related requirements.
        /// </summary>
        private void ValidateViewFields()
        {
            // If ViewFields contains any element, and its schema has been validated before de-serialization; so this requirement MS-WSSCAML_R894 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                "MS-WSSCAML",
                894,
                @"[In Schema] [The schema definition of FieldRefDefinitionView is as follows:] <xs:complexType name=""FieldRefDefinitionView"">
                                                              <xs:attribute name=""Name"" type=""xs:string"" use=""required"" />
                                                              <xs:attribute name=""Explicit"" type=""TRUE_If_Present"" use=""optional"" default=""FALSE""/>
                                                              </xs:complexType>");
        }

        /// <summary>
        /// Used to validate the Brief View Definition related requirement.
        /// </summary>
        /// <param name="view">specify the type of view</param>
        private void ValidateBriefViewDefinition(BriefViewDefinition view)
        {
            // Verify MS-VIEWSS requirement: MS-VIEWSS_R158
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                158,
                @"[In BriefViewDefinition] The definition of the BriefViewDefinition element is as follows.
                                            <s:complexType name=""BriefViewDefinition"" mixed=""true"">
                                                <s:sequence>
                                                <s:element name=""Query"" type=""core:CamlQueryRoot"" minOccurs=""1"" maxOccurs=""1"" />
                                                <s:element name=""ViewFields"" minOccurs=""1"" maxOccurs=""1"">
                                                    <s:complexType>
                                                    <s:sequence>
                                                        <s:element name=""FieldRef"" type=""core:FieldRefDefinitionView"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                                    </s:sequence>
                                                    </s:complexType>
                                                </s:element>
                                                <s:element name=""ViewData"" minOccurs=""0"" maxOccurs=""1"">
                                                    <s:complexType>
                                                    <s:sequence>
                                                        <s:element name=""FieldRef"" type=""core:FieldRefDefinitionViewData"" minOccurs=""3"" maxOccurs=""5"" />
                                                    </s:sequence>
                                                    </s:complexType>
                                                </s:element>
                                                <s:element name=""CalendarViewStyles"" type=""core:CalendarViewStylesDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                <s:element name=""RowLimit"" type=""core:RowLimitDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                <s:element name=""Formats"" type=""core:ViewFormatDefinitions"" minOccurs=""0"" maxOccurs=""1""  />
                                                <s:element name=""Aggregations"" type=""core:AggregationsDefinition"" minOccurs=""0"" maxOccurs=""1"" />
                                                <s:element name=""ViewStyle"" type=""core:ViewStyleReference"" minOccurs=""0"" maxOccurs=""1"" />
                                                <s:element name=""OpenApplicationExtension"" type=""s:string"" minOccurs=""0"" maxOccurs=""1""  />
                                                </s:sequence>
                                                <s:attributeGroup ref=""tns:ViewAttributeGroup""/>
                                            </s:complexType>");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R161
            // In the XSD, the two attributes' definitions are the same, so if the schema validation succeeds, then the requirement R161 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                161,
                @"[In ViewAttributeGroup] All attributes specified by ViewAttributeGroup are the same as those specified by the ViewDefinition complex type as specified in [MS-WSSCAML] section 2.3.2.17.");

            // Validate the query info related requirements.
            this.ValidateQuery(view.Query);

            if (view.RowLimit != null)
            {
                this.ValidateRowLimit();
            }

            this.ValidateFormats(view.Formats);
        }

        /// <summary>
        /// Used to validate the AddView result related requirements.
        /// </summary>
        /// <param name="addViewResult">Specify the result of the AddView operation.</param>
        private void ValidateAddViewResult(AddViewResponseAddViewResult addViewResult)
        {
            // Verify MS-VIEWSS requirement: MS-VIEWSS_R262
            // If the schema validation succeeds, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                262,
                @"[In AddViewSoapOut] The SOAP body contains an AddViewResponse element (section 3.1.4.1.2.3).");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R92
            // Since the operation succeeds without SOAP exception, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                92,
                "[In AddView] The protocol client sends an AddViewSoapIn request message (section 3.1.4.1.1.1), and the protocol server responds with an AddViewSoapOut response message (section 3.1.4.1.1.2).");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R8002
            // Since the operation succeeds without SOAP exception, then this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                8002,
                @"[In AddView] The definition of the AddView operation is as follows.
                                <wsdl:operation name=""AddView"">
                                    <wsdl:input message=""tns:AddViewSoapIn"" />
                                    <wsdl:output message=""tns:AddViewSoapOut"" />
                                </wsdl:operation>");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R164
            // If the schema validation succeeds, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                164,
                @"[In AddViewResponse] The definition of the AddViewResponse element is as follows.
                                                                            <s:element name=""AddViewResponse"">
                                                                              <s:complexType>
                                                                                <s:sequence>
                                                                                  <s:element name=""AddViewResult"" minOccurs=""1"" maxOccurs=""1"">
                                                                                    <s:complexType>
                                                                                      <s:sequence>
                                                                                        <s:element name=""View"" type=""tns:BriefViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                                                                      </s:sequence>
                                                                                    </s:complexType>
                                                                                  </s:element>
                                                                                </s:sequence>
                                                                              </s:complexType>
                                                                            </s:element>");

            // If the schema validation succeeds, then this requirement MS-VIEWSS_R103 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                103,
                @"[In AddViewResponse] The type of the View element[in AddViewResult] is BriefViewDefinition, which is specified in section 2.2.4.1.");

            // Validate BriefViewDefinition related requirements.
            this.ValidateBriefViewDefinition(addViewResult.View);
        }

        /// <summary>
        /// Used to validate the DeleteView result.
        /// </summary>
        private void ValidateDeleteViewResult()
        {
            // Since the viewssProxy.soapAction has been verified to be the SOAP action value of DeleteViewSoapOut in the method ValidateSOAPAction,which means
            // the empty element DeleteViewResponse is returned in the message DeleteViewSoapOut from server, this requirement MS-VIEWSS_R8006 can be captured.
            Site.CaptureRequirement(
                8006,
                @"[In DeleteView] The definition of the DeleteView operation is as follows.
                                    <wsdl:operation name=""DeleteView"">
                                        <wsdl:input message=""tns:DeleteViewSoapIn"" />
                                        <wsdl:output message=""tns:DeleteViewSoapOut"" />
                                    </wsdl:operation>");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R105
            // If the schema validation succeeds, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                105,
                @"[In DeleteView] The protocol client sends a DeleteViewSoapIn request message (section 3.1.4.2.1.1), and the protocol server responds with a DeleteViewSoapOut response message (section 3.1.4.2.1.2).");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R281
            // If the schema validation succeeds, then this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                281,
                @"[In DeleteViewSoapOut] The SOAP body contains a DeleteViewResponse element (section 3.1.4.2.2.2).");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R166
            // Since the DeleteView method has no returned value, it confirms to the schema defined in TD. 
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                166,
                @"[In DeleteViewResponse] The definition of the DeleteViewResponse element is as follows.<s:element name=""DeleteViewResponse"">
                                                  <s:complexType/>
                                                </s:element>");
        }

        /// <summary>
        /// Used to validate the GetView result.
        /// </summary>
        /// <param name="getViewResult">Specify the result of the GetView operation.</param>
        private void ValidateGetViewResult(GetViewResponseGetViewResult getViewResult)
        {
            // Since the operation succeeds without SOAP exception, this requirement MS-VIEWSS_R8007 can be verified directly.
            Site.CaptureRequirement(
                8007,
                @"[In GetView] The definition of the GetView operation is as follows.
                                <wsdl:operation name=""GetView"">
                                    <wsdl:input message=""tns:GetViewSoapIn"" />
                                    <wsdl:output message=""tns:GetViewSoapOut"" />
                                </wsdl:operation>");

            // If the schema validation succeeds, then MS-VIEWSS_R291 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                291,
                @"[In GetViewSoapOut] The SOAP body contains a GetViewResponse element (section 3.1.4.3.2.2).");

            // Since the operation succeeds without SOAP exception, then MS-VIEWSS_R108 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                108,
                @"[In GetView] The protocol client sends a GetViewSoapIn request message (section 3.1.4.3.1.1), and the protocol server responds with a GetViewSoapOut response (section 3.1.4.3.1.2).");

            // If schema has been validated before de-serialization; so this requirement MS-VIEWSS_R110 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                110,
                @"[In GetViewResponse] The type of the View element is BriefViewDefinition, which is specified in section 2.2.4.1.");

            // If the schema validation succeeds, the requirement R168 can be captured. 
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                168,
                @"[In GetViewResponse] The definition of the GetViewResponse element is as follows.
                                                                            <s:element name=""GetViewResponse"">
                                                                              <s:complexType>
                                                                                <s:sequence>
                                                                                  <s:element name=""GetViewResult"" minOccurs=""1"" maxOccurs=""1"">
                                                                                    <s:complexType>
                                                                                      <s:sequence>
                                                                                        <s:element name=""View"" type=""tns:BriefViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                                                                      </s:sequence>
                                                                                    </s:complexType>
                                                                                  </s:element>
                                                                                </s:sequence>
                                                                              </s:complexType>
                                                                            </s:element>");

            // Verify the BriefViewDefinition type.
            this.ValidateBriefViewDefinition(getViewResult.View);
        }

        /// <summary>
        /// Used to validate the GetViewCollection result.
        /// </summary>
        /// <param name="getViewCollectionResult">Specify the result of the GetViewCollection operation.</param>
        private void ValidateGetViewCollectionResult(GetViewCollectionResponseGetViewCollectionResult getViewCollectionResult)
        {
            // Since the operation succeeds without SOAP exception, we can validate this MS-VIEWSS_R8009 requirement directly.
            Site.CaptureRequirement(
                8009,
                @"[In GetViewCollection] The definition of the GetViewCollection operation is as follows.
                                    <wsdl:operation name=""GetViewCollection"">
                                        <wsdl:input message=""tns:GetViewCollectionSoapIn"" />
                                        <wsdl:output message=""tns:GetViewCollectionSoapOut"" />
                                    </wsdl:operation>");

            // Since the operation succeeds without SOAP exception, then the requirement MS-VIEWSS_R112 can be captured directly
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                112,
                @"[In GetViewCollection] The protocol client sends a GetViewCollectionSoapIn request message (section 3.1.4.4.1.1), and the protocol server responds with a GetViewCollectionSoapOut response message (section 3.1.4.4.1.2).");

            // If the schema validation succeeds, then requirement MS-VIEWSS_R299 can be captured. 
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                299,
                @"[In GetViewCollectionSoapOut] The SOAP body contains a GetViewCollectionResponse element (section 3.1.4.4.2.2).");

            // Since the schema validation succeeds then requirement R170 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                170,
                @"[In GetViewCollectionResponse] The definition of the GetViewCollectionResponse element is as follows.
                                                                                     <s:element name=""GetViewCollectionResponse"">
                                                                                     <s:complexType>
                                                                                     <s:sequence>
                                                                                     <s:element name=""GetViewCollectionResult"" minOccurs=""1"" maxOccurs=""1"">
                                                                                     <s:complexType>
                                                                                     <s:sequence>
                                                                                     <s:element name=""Views"" minOccurs=""1"" maxOccurs=""1"">
                                                                                     <s:complexType>
                                                                                     <s:sequence>
                                                                                     <s:element name=""View"" minOccurs=""0"" maxOccurs=""unbounded"">
                                                                                     <s:complexType>
                                                                                     <s:attributeGroup ref=""tns:ViewAttributeGroup""/>
                                                                                     </s:complexType>
                                                                                     </s:element>
                                                                                     </s:sequence>
                                                                                     </s:complexType>
                                                                                     </s:element>
                                                                                     </s:sequence>
                                                                                     </s:complexType>
                                                                                     </s:element>
                                                                                     </s:sequence>
                                                                                     </s:complexType>
                                                                                     </s:element>");

            // Verify MS-VIEWSS requirement: MS-VIEWSS_R161
            // In the XSD, the both attribute definition are the same, so if the schema validation succeeds, then the requirement R161 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                161,
                @"[In ViewAttributeGroup] All attributes specified by ViewAttributeGroup are the same as those specified by the ViewDefinition complex type as specified in [MS-WSSCAML] section 2.3.2.17.");

            foreach (GetViewCollectionResponseGetViewCollectionResultView view in getViewCollectionResult.Views)
            {
                // If the URL is server-relative URL and the schema validation succeeds, then capture requirement R114. 
                Site.CaptureRequirementIfIsTrue(
                    this.IsServerRelativeUrl(view.Url) && this.PassSchemaValidation,
                    114,
                    "[In GetViewCollectionResponse] The attribute group ViewAttributeGroup is specified in section 2.2.8.1, with the following exception: The Url attribute MUST be the server-relative URL of the list view.");
            }
        }

        /// <summary>
        /// Used to validate the GetViewHtml result.
        /// </summary>
        /// <param name="getViewHtmlResult">specify the result of the GetViewHtml operation.</param>
        private void ValidateGetViewHtmlResult(GetViewHtmlResponseGetViewHtmlResult getViewHtmlResult)
        {
            // Since the GetViewHtml operation succeeds without SOAP exception, then the requirement R8010 can be captured directly.
            Site.CaptureRequirement(
                8010,
                "[In GetViewHtml] The definition of the GetViewHtml operation is as follows. <wsdl:operation name=\"GetViewHtml\"> <wsdl:input message=\"tns:GetViewHtmlSoapIn\" /> <wsdl:output message=\"tns:GetViewHtmlSoapOut\" /> </wsdl:operation>");

            // The GetViewHtmlResponse element minOccurs = "1", if the schema validation succeeds, then requirement R309 can be captured. 
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                309,
                @"[In GetViewHtmlSoapOut] The SOAP body contains a GetViewHtmlResponse element (section 3.1.4.5.2.2).");

            // Since the GetViewHtml operation succeeds without SOAP exception, then requirement MS-VIEWSS_R116 can be captured. 
            Site.CaptureRequirement(
                116,
                "[In GetViewHtml] The protocol client sends a GetViewHtmlSoapIn request message (section 3.1.4.5.1.1), and the protocol server responds with a GetViewHtmlSoapOut response message (section 3.1.4.5.1.2).");

            // If the schema validation succeeds, the requirement R172 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                172,
                @"[In GetViewHtmlResponse] The definition of the GetViewHtmlResponse element is as follows.
                                                                                <s:element name=""GetViewHtmlResponse"">
                                                                                  <s:complexType>
                                                                                    <s:sequence>
                                                                                      <s:element name=""GetViewHtmlResult"" minOccurs=""1"" maxOccurs=""1"">
                                                                                        <s:complexType>
                                                                                          <s:sequence>
                                                                                            <s:element name=""View"" type=""core:ViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                                                                          </s:sequence>
                                                                                        </s:complexType>
                                                                                      </s:element>
                                                                                    </s:sequence>
                                                                                  </s:complexType>
                                                                                </s:element>");

            // Validate the view definition related requirements.
            this.ValidateViewDefinition(getViewHtmlResult.View);
        }

        /// <summary>
        /// Used to validate the UpdateView result.
        /// </summary>
        /// <param name="updateViewResult">Specify the result of the UpdateView operation.</param>
        private void ValidateUpdateViewResult(UpdateViewResponseUpdateViewResult updateViewResult)
        {
            // If the UpdateView operation succeeds without SOAP fault exception, the requirement MS-VIEWSS_R8011 can be captured directly.
            Site.CaptureRequirement(
                8011,
                @"[In UpdateView] The definition of the UpdateView element is as follows.
                                    <wsdl:operation name=""UpdateView"">
                                        <wsdl:input message=""tns:UpdateViewSoapIn"" />
                                        <wsdl:output message=""tns:UpdateViewSoapOut"" />
                                    </wsdl:operation>");

            // If the UpdateView operation succeeds without SOAP fault exception, the requirement MS-VIEWSS_R120 can be captured directly.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                120,
                "[In UpdateView] The protocol client sends an UpdateViewSoapIn request message (section 3.1.4.6.1.1), and the protocol server responds with an UpdateViewSoapOut response message (section 3.1.4.6.1.2).");

            // If the schema validation succeeds, then the requirement MS-VIEWSS_R320 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                320,
                @"[In UpdateViewSoapOut] The SOAP body contains an UpdateViewResponse element (section 3.1.4.6.2.2).");

            // If the schema validation succeeds, then the requirement MS-VIEWSS_R174 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                174,
                @"[In UpdateViewResponse] The definition of the UpdateViewResponse element is as follows.
                                <s:element name=""UpdateViewResponse"">
                                <s:complexType>
                                <s:sequence>
                                <s:element name=""UpdateViewResult"" minOccurs=""1"" maxOccurs=""1"">
                                <s:complexType>
                                <s:sequence>
                                <s:element name=""View"" type=""tns:BriefViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                </s:sequence>
                                </s:complexType>
                                </s:element>
                                </s:sequence>
                                </s:complexType>
                                </s:element>");

            // Validate the BriefViewDefinition type
            this.ValidateBriefViewDefinition(updateViewResult.View);
        }

        /// <summary>
        /// Used to validate the UpdateViewHtml result.
        /// </summary>
        /// <param name="updateViewHtmlResult">Specify the result of the UpdateViewHtml operation.</param>
        private void ValidateUpdateViewHtmlResult(UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult)
        {
            // Since the operation succeeds without SOAP exception, then the requirement R8012 can be captured directly
            Site.CaptureRequirement(
                8012,
                @"[In UpdateViewHtml] The definition of the UpdateViewHtml element is as follows.
                                        <wsdl:operation name=""UpdateViewHtml"">
                                            <wsdl:input message=""tns:UpdateViewHtmlSoapIn"" />
                                            <wsdl:output message=""tns:UpdateViewHtmlSoapOut"" />
                                        </wsdl:operation>");

            // If the schema validation succeeds, then requirement R336 can captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                336,
                @"[In UpdateViewHtmlSoapOut] The SOAP body contains an UpdateViewHtmlResponse element (section 3.1.4.7.2.2).");

            // Since the operation succeeds without SOAP exception, then the requirement R124 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                124,
                "[In UpdateViewHtml] The protocol client sends an UpdateViewHtmlSoapIn request message (section 3.1.4.7.1.1), and the protocol server responds with an UpdateViewHtmlSoapOut response message (section 3.1.4.7.1.2).");

            // If the schema validation succeeds, then the requirement R176 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                "MS-VIEWSS",
                176,
                @"[In UpdateViewHtmlResponse] The definition of the UpdateViewHtmlResponse element is as follows.
                                        <s:element name=""UpdateViewHtmlResponse"">
                                        <s:complexType>
                                        <s:sequence>
                                        <s:element name=""UpdateViewHtmlResult"" minOccurs=""1"" maxOccurs=""1"">
                                        <s:complexType>
                                        <s:sequence>
                                        <s:element name=""View"" type=""core:ViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                        </s:sequence>
                                        </s:complexType>
                                        </s:element>
                                        </s:sequence>
                                        </s:complexType>
                                        </s:element>");

            // Validate the view definition related requirements.
            this.ValidateViewDefinition(updateViewHtmlResult.View);
        }

        /// <summary>
        /// Used to validate the UpdateViewHtml2 result.
        /// </summary>
        /// <param name="updateViewHtml2Result">Specify the result of the updateViewHtml2 operation.</param>
        private void ValidateUpdateViewHtml2Result(UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Result)
        {
            // Since the UpdateViewHtml2 operation succeeds without SOAP exception, the requirement R8013 can be captured directly.
            Site.CaptureRequirement(
                8013,
                @"[In UpdateViewHtml2] The definition of the UpdateViewHtml2 operation is as follows.
                                        <wsdl:operation name=""UpdateViewHtml2"">
                                            <wsdl:input message=""tns:UpdateViewHtml2SoapIn"" />
                                            <wsdl:output message=""tns:UpdateViewHtml2SoapOut"" />
                                        </wsdl:operation>");

            // Since the operation succeeds without SOAP exception, then the requirement R130 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                130,
                "[In UpdateViewHtml2] The protocol client sends an UpdateViewHtml2SoapIn request message (section 3.1.4.8.1.1), and the protocol server responds with an UpdateViewHtml2SoapOut response message (section 3.1.4.8.1.2).");

            // The UpdateViewHtml2Response minOccurs="1", the schema validation succeeds then requirement R344 can captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                344,
                @"[In UpdateViewHtml2SoapOut] The SOAP body contains an UpdateViewHtml2Response element (section 3.1.4.8.2.8).");

            // If the schema validation succeeds then requirement R184 can be captured.
            Site.CaptureRequirementIfIsTrue(
                this.PassSchemaValidation,
                184,
                @"[In UpdateViewHtml2Response] The definition of the UpdateViewHtml2Response element is as follows.
                            <s:element name=""UpdateViewHtml2Response"">
                              <s:complexType>
                                <s:sequence>
                                  <s:element name=""UpdateViewHtml2Result"" minOccurs=""1"" maxOccurs=""1"">
                                    <s:complexType>
                                      <s:sequence>
                                        <s:element name=""View"" type=""core:ViewDefinition"" minOccurs=""1"" maxOccurs=""1""/>
                                      </s:sequence>
                                    </s:complexType>
                                  </s:element>
                                </s:sequence>
                              </s:complexType>
                            </s:element>");

            // Validate the view definition related requirements.
            this.ValidateViewDefinition(updateViewHtml2Result.View);
        }
    }
}