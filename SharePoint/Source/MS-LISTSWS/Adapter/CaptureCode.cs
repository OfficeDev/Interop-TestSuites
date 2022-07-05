namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial class is the container for the adapter capture code.
    /// </summary>
    public partial class MS_LISTSWSAdapter : ManagedAdapterBase, IMS_LISTSWSAdapter
    {
        /// <summary>
        /// Verify the requirements of the transport when the response is received successfully.
        /// </summary>
        private void VerifyTransportRequirements()
        {
            string transport = Common.GetConfigurationPropertyValue("TransportType", this.Site);

            if (string.Compare(transport, AdapterHelper.TransportHttp, StringComparison.OrdinalIgnoreCase) == 0)
            {
                // Verify MS-LISTSWS requirement MS-LISTSWS_R1.
                // Having received the response successfully have proved the HTTP 
                // transport is supported. If the HTTP transport is not supported, the response 
                Site.CaptureRequirement(
                    1,
                    @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }
            else
            {
                if (string.Compare(transport, AdapterHelper.TransportHttps, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    if (Common.IsRequirementEnabled(3, this.Site))
                    {
                        // Verify MS-LISTSWS requirement: MS-LISTSWS_R3
                        // Having received the response successfully have proved the HTTPS 
                        // transport is supported. If the HTTPS transport is not supported, the 
                        // response can't be received successfully.
                        Site.CaptureRequirement(
                            3,
                            @"[In Transport] Implementation does additionally support SOAP over HTTPS for securing communication with protocol clients. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
                    }
                }
                else
                {
                    this.Site.Assert.Fail("The transport is not valid for the value {0}", transport);
                }
            }

            SoapProtocolVersion soapVersion = this.listsProxy.SoapVersion;

            if (soapVersion == SoapProtocolVersion.Soap11
                || soapVersion == SoapProtocolVersion.Soap12)
            {
                // Verify MS-LISTSWS requirement MS-LISTSWS_R4
                // Having received the response successfully have proved the message 
                // format is correct. If the message format is incorrect, the response can't be
                // received successfully.
                Site.CaptureRequirement(
                    4,
                    @"[In Transport]Protocol messages MUST be formatted as specified either in "
                    + @"[SOAP1.1], ""SOAP Envelope"", or in [SOAP1.2-1/2017], ""SOAP Message Construct"".");
            }

            // Verify R1177
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured. 
            Site.CaptureRequirement(
                1177,
                @"[In Common Message Syntax]The syntax of the definitions uses XML schema, "
                + "as specified in [XMLSCHEMA1/2] and [XMLSCHEMA2/2], and WSDL, as specified in [WSDL].");

            // Verify R1178
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1178,
                @"[In Namespaces]This protocol specifies and references XML namespaces by "
                + "using the mechanisms specified in [XMLNS].");
        }

        /// <summary>
        /// Verify the soap exception fault.
        /// </summary>
        private void VerifySoapExceptionFault()
        {
            // COMMENT: If the SOAP error code is not equal to OperationSucceedErrorCode, 
            // which means an exception is thrown, then the following requirement can be 
            // captured.
            Site.CaptureRequirement(
                7,
                @"[In Transport]Protocol server faults MUST be returned either by using HTTP "
                + @"Status Codes as specified in [RFC2616], section 10, ""Status Code Definitions"", "
                + "or by using SOAP faults as specified either in [SOAP1.1], section 4.4, "
                + @"""SOAP Fault"", or in [SOAP1.2-1/2017], section 5.4, ""SOAP Fault"".");
        }

        #region Capture Adapter requirements of Complex Types

        /// <summary>
        /// Verify the requirements of the complex type CamlQueryOptions.
        /// </summary>
        /// <param name="query">The actual CamlQueryOptions.</param>
        /// <param name="viewFields">Specifies which fields of the list item should be returned</param>
        /// <param name="returnedAuthorField">The actual returned Author field.</param>
        private void VerifyCamlQueryOptions(
            CamlQueryOptions query,
            CamlViewFields viewFields,
            string returnedAuthorField)
        {
            Site.Assert.IsNotNull(query, "The CamlQueryOptions cannot be null");

            // Verify R34
            bool authorContained = false;
            if (viewFields != null)
            {
                if (viewFields.ViewFields != null)
                {
                    if (viewFields.ViewFields.FieldRef != null)
                    {
                        foreach (CamlViewFieldsViewFieldsFieldRef fr in viewFields.ViewFields.FieldRef)
                        {
                            if (string.Compare(fr.Name, AdapterHelper.FieldAuthorName, StringComparison.OrdinalIgnoreCase) == 0)
                            {
                                authorContained = true;
                                break;
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(query.QueryOptions.ExpandUserField))
            {
                if (bool.Parse(query.QueryOptions.ExpandUserField) && authorContained)
                {
                    // Verify MS-LISTSWS requirement: MS-LISTSWS_R34
                    // If the returned Author field contains ",#", which means "Name", 
                    // "EMail", "SipAddress", and "Title" fields from the user information List are 
                    // returned, then the requirement can be captured.
                    bool isVerifyR34 = false;
                    if (!string.IsNullOrEmpty(returnedAuthorField))
                    {
                        isVerifyR34 = returnedAuthorField.Contains(",#");
                    }

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR34,
                        34,
                        @"[In CAMLQueryOptions] [ExpandUserField] If set to True, specifies that fields "
                        + "in list items that are lookup fields to the user information list are returned as "
                        + @"if they were multi-value lookups, including ""Name"", ""EMail"", "
                        + @"""SipAddress"", and ""Title"" fields from the user information List for the "
                        + "looked-up item.");
                }
            }
        }

        /// <summary>
        /// Verify the requirements of the complex type DataDefinition.
        /// </summary>
        /// <param name="data">The actual DataDefinition complex type.</param>
        private void VerifyDataDefinition(DataDefinition data)
        {
            // Verify R1296
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1296,
                @"[DataDefinition]Specifies items contained within a list.");

            // Verify R1297
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1297,
                @"[The schema of DataDefinition is defined as:] "
                + @"<s:complexType name=""DataDefinition"" mixed=""true"">"
                + @"  <s:sequence>"
                + @"    <s:any minOccurs=""0"" maxOccurs=""unbounded""/>"
                + @"  </s:sequence>"
                + @" <s:attribute name=""ItemCount"" type=""s:string"" use=""required"" />"
                + @" <s:attribute name=""ListItemCollectionPositionNext"" type =""s:string"" use=""optional""/>"
                + @"</s:complexType>");

            Site.Assert.IsNotNull(data, "The data cannot be null");

            // Verify R1298
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1298,
                @"[DataDefinition]The DataDefinition element contains a required ItemCount "
                + "attribute and an optional ListItemCollectionPositionNext attribute.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1299
            // If the ItemCount can be parsed to a UInt32 and the value equal to the
            // count of the row in the element, then the following requirement can be captured.
            uint itemCountValue = 0;
            this.Site.Assert.IsTrue(
                uint.TryParse(data.ItemCount, out itemCountValue),
                "The value of ItemCount[{0}] should be a valid uint32 value.",
                string.IsNullOrEmpty(data.ItemCount) ? "Null" : data.ItemCount);

            if (itemCountValue.Equals(0))
            {
                // if the item count equal to 0, there should be no any row data, then capture R1299
                Site.CaptureRequirementIfIsNull(
                 data.Any,
                 1299,
                 @"[DataDefinition]The ItemCount attribute is an unsigned 32-bit integer that "
                 + "specifies the number of list items that are included in the response.");
            }
            else
            {
                // if the item count does not equal to 0, the number of row items should be equal to the itemCount, then capture R1299
                this.Site.Assert.IsNotNull(data.Any, "The row data should not be null, if the item[{0}] does not equal to Zero", itemCountValue);
                Site.CaptureRequirementIfAreEqual(
                 itemCountValue,
                 (uint)data.Any.Length,
                 1299,
                 @"[DataDefinition]The ItemCount attribute is an unsigned 32-bit integer that "
                 + "specifies the number of list items that are included in the response.");
            }

            // Verify R1300
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1300,
                @"[DataDefinition]ListItemCollectionPositionNext is used "
                + "by protocol server methods that support paging of results "
                + "and it[ListItemCollectionPositionNext] is an opaque string returned by the protocol server "
                + "that allows the protocol client to pass in a subsequent call to get the next page of data.");

            // Verify R1301
            if (int.Parse(data.ItemCount) > 0)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1301
                Site.CaptureRequirementIfAreEqual<int>(
                    int.Parse(data.ItemCount),
                    data.Any.Length,
                    1301,
                    @"[DataDefinition]The DataDefinition element further contains a number of z:row "
                    + "elements equal to the value of the ItemCount attribute.");
            }

            // Verify R2392
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2392,
                @"data: Within the data element, zero or more 'z:row' sub-elements are "
                + @"present, denoting ""rows"" of the tabular data.");

            // Verify R2255
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2255,
                @"[DataDefinition]Each z:row element describes a single list item.");

            // Verify R1302
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1302,
                @"[DataDefinition]z is equal to #RowsetSchema in the Microsoft ADO 2.6 "
                + "Persistence format (as specified in  [MS-PRSTFR]).");

            // Verify R1202
            // If all the above requirements are verified, then the requirement can be 
            // captured.
            Site.CaptureRequirement(
                1202,
                @"[In Complex Types]The Complex type DataDefinition is used Specifies items "
                + "contained within a list.");
        }

        /// <summary>
        /// Verify the requirements of the complex type FieldReferenceDefinitionCT.
        /// </summary>
        /// <param name="field">The actual FieldReferenceDefinitionCT complex type.</param>
        private void VerifyFieldReferenceDefinitionCT(
            FieldReferenceDefinitionCT field)
        {
            // Verify R1314
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1314,
                @"[FieldReferenceDefinitionCT]Specifies data on a field included in a content type.");

            // Verify R1315
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1315,
                @"[The schema of FieldReferenceDefinitionCT is defined as:]"
                + @"<s:complexType name=""FieldReferenceDefinitionCT"">"
                + @"  <s:attribute name=""Aggregation"" type=""s:string""/>"
                + @"  <s:attribute name=""Customization"" type=""s:string""/>"
                + @"  <s:attribute name=""DisplayName"" type=""s:string""/>"
                + @"  <s:attribute name=""Format"" type=""s:string""/>"
                + @"  <s:attribute name=""Hidden"" type=""core:TRUEFALSE""/>"
                + @"  <s:attribute name=""ID"" type=""core:UniqueIdentifierWithOrWithoutBraces"" "
                + @"               use=""required""/>"
                + @"  <s:attribute name=""Name"" type=""s:string"" use=""required""/>"
                + @"  <s:attribute name=""Node"" type=""s:string""/>"
                + @"  <s:attribute name=""PIAttribute"" type=""s:string""/>"
                + @"  <s:attribute name=""PITarget"" type=""s:string""/>"
                + @"  <s:attribute name=""PrimaryPIAttribute"" type=""s:string""/>"
                + @"  <s:attribute name=""PrimaryPITarget"" type=""s:string""/>"
                + @"  <s:attribute name=""ReadOnly"" type=""core:TRUEFALSE""/>"
                + @"  <s:attribute name=""Required"" type=""core:TRUEFALSE""/>"
                + @"  <s:attribute name=""ShowInEditForm"" type=""core:TRUEFALSE""/>"
                + @"  <s:attribute name=""ShowInNewForm"" type=""core:TRUEFALSE""/>"
                + @"</s:complexType>");

            Site.Assert.IsNotNull(field, "The FieldReferenceDefinitionCT cannot be null");

            // Verify R1316
            if (!string.IsNullOrEmpty(field.Node))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1316
                // For the fields with the Node or node attribute, if the Aggregation 
                // attribute is not null, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    field.Aggregation,
                    1316,
                    @"[FieldReferenceDefinitionCT.Aggregation:] Overrides the field value specified in [MS-WSSFO2] section 2.2.8.3.3.2 for this reference to the field.
                    [Aggregation: For fields with the Node or node attribute, a reader MUST use this attribute[Aggregation] to control promotion from XML or site template files.
                    For other fields, a reader MUST ignore this attribute[Aggregation].]");
            }

            // Verify R1205
            // If all the above requirements are verified, then the requirement can be 
            // captured.
            Site.CaptureRequirement(
                1205,
                @"[In Complex Types]The Complex type FieldReferenceDefinitionCT is used Specifies data "
                + "on a field included in a content type.");
        }

        /// <summary>
        /// Verify the requirements of the complex type FileFolderChangeDefinition.
        /// </summary>
        /// <param name="change">The actual FileFolderChangeDefinition complex type.</param>
        private void VerifyFileFolderChangeDefinition(
            FileFolderChangeDefinition change)
        {
            Site.Assert.IsNotNull(change, "The FileFolderChangeDefinition cannot be null");

            // Verify the requirements of the ChangeTypeEnum simple type.
            this.VerifyChangeTypeEnum();
        }

        /// <summary>
        /// Verify the requirements of the complex type FileFragmentChangeDefinition.
        /// </summary>
        /// <param name="change">The actual FileFragmentChangeDefinition complex type.</param>
        private void VerifyFileFragmentChangeDefinition(
            FileFragmentChangeDefinition change)
        {
            if (change != null)
            {
                // Verify the requirements of the ChangeTypeEnum simple type.
                this.VerifyChangeTypeEnum();
            }
        }

        /// <summary>
        /// Verify the requirements of the complex type ListDefinitionCT.
        /// </summary>
        /// <param name="list">The actual ListDefinitionSchema which is an extension based on ListDefinitionCT.</param>
        private void VerifyListDefinitionCT(ListDefinitionCT list)
        {
            Site.Assert.IsNotNull(list, "The ListDefinitionCT cannot be null");

            // Verify R1347
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1347,
                @"[ ListDefinitionCT]Specifies information about a particular list.");

            // Verify R1348
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1348,
               @"[The schema of ListDefinitionCT is defined as:]<s:complexType name=""ListDefinitionCT"">
                  <s:complexType name=""ListDefinitionCT"">
                  <s:attribute name=""DocTemplateUrl"" type=""s:string"" />
                  <s:attribute name=""DefaultViewUrl"" type=""s:string"" use=""required"" />
                  <s:attribute name=""MobileDefaultViewUrl"" type=""s:string"" />
                  <s:attribute name=""ID"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Title"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Description"" type=""s:string"" />
                  <s:attribute name=""ImageUrl"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Name"" type=""s:string"" use=""required"" />
                  <s:attribute name=""BaseType"" type=""s:string"" use=""required"" />
                  <s:attribute name=""FeatureId"" type=""s:string"" use=""required"" />
                  <s:attribute name=""ServerTemplate"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Created"" type=""s:string"" use=""required"" />
                  <s:attribute name=""Modified"" type=""s:string"" use=""required"" />
                  <s:attribute name=""LastDeleted"" type=""s:string"" />
                  <s:attribute name=""Version"" type=""s:int"" use=""required"" />
                  <s:attribute name=""Direction"" type=""s:string"" use=""required"" />
                  <s:attribute name=""ThumbnailSize"" type=""s:string"" />
                  <s:attribute name=""WebImageWidth"" type=""s:string"" />
                  <s:attribute name=""WebImageHeight"" type=""s:string"" />
                  <s:attribute name=""Flags"" type=""s:int"" />
                  <s:attribute name=""ItemCount"" type=""s:int"" use=""required"" />
                  <s:attribute name=""AnonymousPermMask"" type=""s:unsignedLong"" />
                  <s:attribute name=""RootFolder"" type=""s:string"" />
                  <s:attribute name=""ReadSecurity"" type=""s:int"" use=""required"" />
                  <s:attribute name=""WriteSecurity"" type=""s:int"" use=""required"" />
                  <s:attribute name=""Author"" type=""s:string"" />
                  <s:attribute name=""EventSinkAssembly"" type=""s:string"" />
                  <s:attribute name=""EventSinkClass"" type=""s:string"" />
                  <s:attribute name=""EventSinkData"" type=""s:string"" />
                  <s:attribute name=""EmailInsertsFolder"" type=""s:string"" />
                  <s:attribute name=""EmailAlias"" type=""s:string"" />
                  <s:attribute name=""WebFullUrl"" type=""s:string"" />
                  <s:attribute name=""WebId"" type=""s:string"" />
                  <s:attribute name=""SendToLocation"" type=""s:string"" />
                  <s:attribute name=""ScopeId"" type=""s:string"" />
                  <s:attribute name=""MajorVersionLimit"" type=""s:int"" />
                  <s:attribute name=""MajorWithMinorVersionsLimit"" type=""s:int"" />
                  <s:attribute name=""WorkFlowId"" type=""s:string"" />
                  <s:attribute name=""HasUniqueScopes"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""NoThrottleListOperations"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""HasRelatedLists"" type=""s:string"" />
                  <s:attribute name=""AllowDeletion"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""AllowMultiResponses"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableAttachments"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableModeration"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableVersioning"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""HasExternalDataSource"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""Hidden"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""MultipleDataList"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""Ordered"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""ShowUser"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnablePeopleSelector"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableResourceSelector"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableMinorVersion"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""RequireCheckout"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""ThrottleListOperations"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""ExcludeFromOfflineClient"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""CanOpenFileAsync"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnableFolderCreation"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""IrmEnabled"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""IsApplicationList"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""PreserveEmptyValues"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""StrictTypeCoercion"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""EnforceDataValidation"" type=""core:TRUEFALSE"" use=""optional""/>
                  <s:attribute name=""MaxItemsPerThrottledOperation"" type=""s:int"" />
                  <s:attribute name=""EnableAssignedToEmail"" type=""core:TRUEFALSE"" use=""optional""/>
                  <s:attribute name=""Followable"" type=""core:TRUEFALSE"" />
                  <s:attribute name=""Acl"" type =""s: string"" use =""optional""/>
                  <s:attribute name=""Flags2"" type = ""s:int"" use = ""optional""/>
	              <s:attribute name=""ComplianceTag"" type=""s:string"" use=""optional""/>
	              <s:attribute name=""ComplianceFlags"" type=""s:int"" use=""optional""/>
	              <s:attribute name=""UserModified"" type=""s:dateTime"" use=""optional""/>
	              <s:attribute name=""ListSchemaVersion"" type=""s:int"" use=""optional""/>
	              <s:attribute name=""AclVersion"" type=""s:int"" use=""optional""/>
                  <s:attribute name=""RootFolderId"" type = ""s:string"" use = ""optional""/>
                  <s:attribute name=""IrmSyncable"" type = ""core:TRUEFALSE"" use = ""optional""/>
                  </s:complexType> ");

            if (Common.IsRequirementEnabled(5417, this.Site))
            {
                this.Site.CaptureRequirementIfIsNull(
                  list.Followable,
                  5417,
                  @"Implementation does not return this attribute[ListDefinitionCT.Followable]. [In Appendix B: Product Behavior] <16> Section 2.2.4.11: This attribute[ListDefinitionCT.Followable] is not returned in Windows SharePoint Services 3.0 and SharePoint Foundation 2010.");
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1352
            // If the ID can be used to create a Guid, then the following requirement 
            // can be captured.
            Guid id;
            bool isVerifyR1352 = Guid.TryParse(list.ID, out id);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1352,
                1352,
                @"[ListDefinitionCT.ID is ]The GUID[ for the list.]");

            Guid guidOfList = new Guid();
            if (!string.IsNullOrEmpty(list.FeatureId))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2236
                Site.CaptureRequirementIfIsTrue(
                    Guid.TryParse(list.FeatureId, out guidOfList),
                    2236,
                    @"[ListDefinitionCT.ID:] The GUID for the list.");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2238
                // If the FeatureID can be used to create a Guid, then the following 
                // requirement can be captured.
                bool isVerifyR2238 = Guid.TryParse(list.FeatureId, out id);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2238,
                    2238,
                    @"[ListDefinitionCT.FeatureID is] The GUID [of the feature that contains the list "
                    + "schema for the list]");
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1356
            // If the Name can be used to create a Guid, then the following requirement 
            // can be captured.
            bool isVerifyR1356 = Guid.TryParse(list.Name, out id);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1356,
                1356,
                @"[ListDefinitionCT.Name is ]The internal name for the list.");

            //Verify MS-LISTSWS requirement: MS-LISTSWS_R1356001
            Site.Assert.IsTrue(list.Name == list.ID, "The Name is equal to ID.");
            Site.CaptureRequirement(
                1356001,
                @"[ListDefinitionCT.Name] The Name is equal to ID.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R52.
            // If the actual BaseType value is contained in the expected domain of 
            // values, then the requirement can be captured.
            string[] baseTypeDomain = { "-1", "0", "1", "3", "4", "5" };

            Site.CaptureRequirementIfIsTrue(
                new List<string>(baseTypeDomain).Contains(list.BaseType),
                52,
                @"[ListDefinitionCT.BaseType] See [MS-WSSFO2] section 2.2.3.11 for the possible "
                + "values of the BaseType."
                + "[The only valid values of the List Base Type are specified as follows."
                + "    "
                + "Value  Description"
                + "0      Generic list"
                + "1      Document library"
                + "3      Discussion board list"
                + "4      Survey list"
                + "5      Issues list]");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1358
            Guid guidOfFeature = new Guid();
            Site.CaptureRequirementIfIsTrue(
                Guid.TryParse(list.ID, out guidOfFeature),
                1358,
                @"[ListDefinitionCT.FeatureID] The GUID of the feature that contains the list "
                + "schema for the list.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1361
            // If the Created can be parsed to a DateTime, then the following 
            // requirement can be captured.
            DateTime created;
            string parseFormat = @"yyyyMMdd HH:mm:ss";
            bool isVerifyR1361 = DateTime.TryParseExact(list.Created, parseFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out created);

            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual values: list.Created[{0}] for requirement #R1361",
                string.IsNullOrEmpty(list.Created) ? "NullOrEmpty" : list.Created);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1361,
                1361,
                @"ListDefinitionCT.Created: Specifies the Coordinated Universal Time (UTC) date "
                + "and time in the Gregorian calendar when the list was created in the format "
                + @"""yyyyMMdd hh:mm:ss"" where ""yyyy"" represents the year, ""MM"" "
                + @"represents the month, ""dd"" represents the day of the month, ""hh"" "
                + @"represents the hour, ""mm"" represents the minute, and ""ss"" represents the "
                + "second.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1362
            // If the Modified can be parsed to a DateTime, then the following 
            // requirement can be captured.
            bool isVerifyR1362 = DateTime.TryParseExact(list.Created, parseFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out created);

            Site.Log.Add(
             LogEntryKind.Debug,
             "The actual values: list.Created[{0}] for requirement #R1362",
             string.IsNullOrEmpty(list.Created) ? "NullOrEmpty" : list.Created);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1362,
                1362,
                @"[ListDefinitionCT.Modified: ]Specifies the Coordinated Universal Time (UTC) date "
                + "and time in the Gregorian calendar when the list was last modified in the format "
                + @"""yyyyMMdd hh:mm:ss"" where ""yyyy"" represents the year, ""MM"" represents "
                + @"the month, ""dd"" represents the day of the month, ""hh"" represents the hour, "
                + @"""mm"" represents the minute, and ""ss"" represents the second.");

            // Verify R1363
            if (!string.IsNullOrEmpty(list.LastDeleted))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1363
                // If the LastDeleted can be parsed to a DateTime, then the following 
                // requirement can be captured.
                bool isR1363 = DateTime.TryParseExact(list.Created, parseFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out created);

                Site.Log.Add(
                 LogEntryKind.Debug,
                 "The actual values: list.Created[{0}] for requirement #R1363",
                 string.IsNullOrEmpty(list.Created) ? "NullOrEmpty" : list.Created);

                Site.CaptureRequirementIfIsTrue(
                    isR1363,
                    1363,
                    @"[ListDefinitionCT.LastDeleted: ]Specifies the Coordinated Universal Time (UTC) "
                    + "date and time in the Gregorian calendar when the list last had an element "
                    + @"deleted in the format ""yyyyMMdd hh:mm:ss"" where""'yyyy"" represents the "
                    + @"year, ""MM"" represents the month, ""dd"" represents the day of the month, "
                    + @"""hh"" represents the hour, ""mm"" represents the minute, and ""ss"" "
                    + "represents the second.");
            }

            // Verify R1365
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1365,
                @"[ListDefinitionCT.Direction: ]Specifies the direction that items in the list are laid "
                + "out in when displayed.");

            if (Common.IsRequirementEnabled(1518047, this.Site))
            {
                // Verify R542000801
                Site.CaptureRequirementIfIsNotNull(
                    list.ComplianceTag,
                    542000801,
                    @"[ListDefinitionCT.ComplianceTag:]Specifies compliance tag.");

                // Verify R542000802
                Site.CaptureRequirementIfIsNotNull(
                    list.ComplianceFlags,
                    542000802,
                    @"[ListDefinitionCT.ComplianceFlags:]Specifies compliance flags.");

                // Verify R542000803
                // If the UserModified can be parsed to a DateTime, then the following 
                // requirement can be captured.
                bool isR542000803 = DateTime.TryParseExact(list.UserModified, parseFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out created);
                Site.CaptureRequirementIfIsTrue(
                    isR542000803,
                    542000803,
                    @"[ListDefinitionCT.UserModified:]Specifies the date and time.");

                // Verify R542000804
                Site.CaptureRequirementIfIsNotNull(
                    list.ListSchemaVersion,
                    542000804,
                    @"[ListDefinitionCT.ListSchemaVersion:]Specifies the version for list schema.");

                // Verify R542000805
                Site.CaptureRequirementIfIsNotNull(
                    list.AclVersion,
                    542000805,
                    @"[ListDefinitionCT.AclVersion:]Specifies the version for access control list (ACL).");
            }            

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R121.
            // If the actual Direction value is contained in the expected domain of 
            // values, then the requirement can be captured.
            string[] directionDomain = { "none", "ltr", "rtl" };
            Site.CaptureRequirementIfIsTrue(
                new List<string>(directionDomain).Contains(list.Direction),
                121,
                @"[ListDefinitionCT.Direction] MUST be one of the following values: none ltr, rtl.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R141.
            // If the actual ReadSecurity value is contained in the expected domain of 
            // values, then the requirement can be captured.
            bool isVerifyR141 = (list.ReadSecurity == 1) || (list.ReadSecurity == 2);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR141,
                141,
                @"[ListDefinitionCT.ReadSecurity] The read permission setting for this list MUST "
                + "be one of the following values: 1, 2.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R144.
            // If the actual WriteSecurity value is contained in the expected domain of 
            // values, then the requirement can be captured.
            bool isVerifyR144 = (list.WriteSecurity == 1) || (list.WriteSecurity == 2)
                || (list.WriteSecurity == 4);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR144,
                144,
                @"[ListDefinitionCT.WriteSecurity] The write permission setting for this list MUST "
                + "be one of the following values: 1, 2, 4.");

            // Verify R2239
            if (!string.IsNullOrEmpty(list.WebId))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2239
                // If the WebId can be used to create a Guid, then the following 
                // requirement can be captured.
                bool isVerifyR2239 = Guid.TryParse(list.WebId, out id);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2239,
                    2239,
                    @"[ListDefinitionCT.WebId is] The GUID [of the site that this list is associated with.]");
            }

            // Verify R2240
            if (!string.IsNullOrEmpty(list.ScopeId))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2240
                // If the ScopeId can be used to create a Guid, then the following 
                // requirement can be captured.
                bool isVerifyR2240 = Guid.TryParse(list.ScopeId, out id);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2240,
                    2240,
                    @"[ListDefinitionCT.ScopeId is]The GUID [of the site that contains this list]");
            }

            // Verify R2241
            if (!string.IsNullOrEmpty(list.WorkFlowId))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2241
                // If the ScopeId can be used to create a Guid, then the following 
                // requirement can be captured.
                bool isVerifyR2241 = Guid.TryParse(list.WorkFlowId, out id);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2241,
                    2241,
                    @"[ListDefinitionCT.WorkFlowId is] The GUID [of a workflow association that is used "
                    + "to manage the Content Approval process for the list.]");
            }

            // Verify R1208
            // If all the above requirements are verified, then the requirement can be 
            // captured.
            Site.CaptureRequirement(
                1208,
                @"[In Complex Types]The Complex type ListDefinitionCT is used Specifies information "
                + "about a particular list.");

            if (Common.IsRequirementEnabled(2396, this.Site))
            {
                // Verify R2396
                Site.CaptureRequirementIfIsNull(
                    list.HasExternalDataSource,
                    2396,
                    @"Implementation does not return this attribute[ListDefinitionCT.HasExternalDataSource]. [In Appendix B: Product Behavior] <4> Section 2.2.4.11: This attribute[ListDefinitionCT.HasExternalDataSource] is not returned by Windows SharePoint Services 3.0 servers.");
            }

            if (Common.IsRequirementEnabled(2398, this.Site))
            {
                // Verify R2398
                Site.CaptureRequirementIfIsNull(
                    list.EnablePeopleSelector,
                    2398,
                    @"Implementation does not return this attribute[ListDefinitionCT.EnablePeopleSelector]. [In Appendix B: Product Behavior] <5> Section 2.2.4.11: This attribute[ListDefinitionCT.EnablePeopleSelector] is not returned by Windows SharePoint Services 3.0 servers.");
            }

            if (Common.IsRequirementEnabled(2400, this.Site))
            {
                // Verify R2400
                Site.CaptureRequirementIfIsNull(
                    list.EnableResourceSelector,
                    2400,
                    @"Implementation does not return this attribute[ListDefinitionCT.EnableResourceSelector]. [In Appendix B: Product Behavior] <6> Section 2.2.4.11: This attribute[ListDefinitionCT.EnableResourceSelector] is not returned by Windows SharePoint Services 3.0 servers.");
            }

            if (Common.IsRequirementEnabled(2402, this.Site))
            {
                // Verify R2402
                Site.CaptureRequirementIfIsNull(
                    list.ExcludeFromOfflineClient,
                    2402,
                    @"Implementation does not return this attribute[ListDefinitionCT.ExcludeFromOfflineClient]. [In Appendix B: Product Behavior] <7> Section 2.2.4.11: This attribute[ListDefinitionCT.ExcludeFromOfflineClient] is not returned by Windows SharePoint Services 3.0 servers.");
            }
            if (Common.IsRequirementEnabled(1401002002, this.Site))
            { 
                //Verify 1401002002
                Site.CaptureRequirementIfIsNull(
                    list.CanOpenFileAsync,
                    1401002002,
                    @"Implementation does not return to the client, when the client attempts to open files asynchronously from the server. (<8> Section 2.2.4.11:  This attribute is not returned by Windows SharePoint Services 2.0, Windows SharePoint Services 3.0 and SharePoint Foundation 2010.)");
            }
            if (Common.IsRequirementEnabled(1401002001, this.Site))
            {
                //Verify 1401002001
                Site.CaptureRequirementIfIsNotNull(
                    list.CanOpenFileAsync,
                    1401002001,
                    @"Implementation does return to the client, when the client attempts to open files asynchronously from the server. (SharePoint Foundation 2013 and above follow this behavior.)");

                //Verify requirement: MS-LISTSWS_R1401001
                Site.CaptureRequirementIfIsNotNull
                    (
                    list.CanOpenFileAsync,
                    1401001,
                    @"[ListDefinitionCT.CanOpenFileAsync:] True, if the client attempts to open files asynchronously from the server.");

            }
            if (Common.IsRequirementEnabled(2404, this.Site))
            {
                // Verify R2404
                Site.CaptureRequirementIfIsNull(
                    list.EnableFolderCreation,
                    2404,
                    @"Implementation does not return this attribute[ListDefinitionCT.EnableFolderCreation]. [In Appendix B: Product Behavior]<9> Section 2.2.4.11: This attribute is not returned by Windows SharePoint Services 3.0 servers.");
            }

            if (Common.IsRequirementEnabled(2406, this.Site))
            {
                // Verify R2406
                Site.CaptureRequirementIfIsNull(
                    list.IrmEnabled,
                    2406,
                    @"Implementation does not return this attribute[ListDefinitionCT.IrmEnabled]. [In Appendix B: Product Behavior] <10> Section 2.2.4.11: This attribute[ListDefinitionCT.IrmEnabled] is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2408, this.Site))
            {
                // Verify R2408
                Site.CaptureRequirementIfIsNull(
                    list.IsApplicationList,
                    2408,
                    @"Implementation does not return this attribute[ListDefinitionCT.IsApplicationList]. [In Appendix B: Product Behavior] <11> Section 2.2.4.11: This attribute[ListDefinitionCT.IsApplicationList] is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2410, this.Site))
            {
                // Verify R2410
                Site.CaptureRequirementIfIsNull(
                    list.PreserveEmptyValues,
                    2410,
                    @"Implementation does not return this attribute[ListDefinitionCT.PreserveEmptyValues]. [In Appendix B: Product Behavior] <12> Section 2.2.4.11: This attribute[ListDefinitionCT.PreserveEmptyValues] is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2412, this.Site))
            {
                // Verify R2412
                Site.CaptureRequirementIfIsNull(
                    list.StrictTypeCoercion,
                    2412,
                    @"Implementation does not return this attribute[ListDefinitionCT.StrictTypeCoercion]. [In Appendix B: Product Behavior] <13> Section 2.2.4.11: This attribute[ListDefinitionCT.StrictTypeCoercion] is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2414, this.Site))
            {
                // Verify R2414
                Site.CaptureRequirementIfIsFalse(
                    list.MaxItemsPerThrottledOperationSpecified,
                    2414,
                    @"Implementation does not return this attribute[ListDefinitionCT.MaxItemsPerThrottledOperation]. [In Appendix B: Product Behavior] <14> Section 2.2.4.11: This attribute[ListDefinitionCT.MaxItemsPerThrottledOperation] is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2416, this.Site))
            {
                // Verify R2416
                Site.CaptureRequirementIfIsNull(
                    list.EnforceDataValidation,
                    2416,
                    @"Implementation does not return this attribute[ListDefinitionCT.EnforceDataValidation]. [In Appendix B: Product Behavior] <15> Section 2.2.4.11: This attribute[ListDefinitionCT.EnforceDataValidation] is not returned in Windows SharePoint Services 3.0.");
            }

            if(Common.IsRequirementEnabled(542000101, this.Site))
            {
                // Verify R542000101
                Site.CaptureRequirementIfIsNotNull(
                    list.Acl,
                    542000101,
                    @"Implementation does return this element[ListDefinitionCT.Acl]. [In Appendix B: Product Behavior] (SharePoint Server 2016 and above support this behavior.)");
            }

            if (Common.IsRequirementEnabled(542000201, this.Site))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R5420003.
                // If the actual BaseType value is contained in the expected domain of 
                // values, then the requirement can be captured.
                string[] Flags2 = { "0", "0x0000000000000001", "0x0000000000000002", "0x0000000000000004", "0x0000000000000008" };

                Site.CaptureRequirementIfIsTrue(
                    new List<string>(Flags2).Contains(list.Flags2),
                    5420003,
                    @"[ListDefinitionCT.Flags2:]This element MUST be one of the following values:[0, 0x0000000000000001, 0x0000000000000002, 0x0000000000000004, 0x0000000000000008]");

                // Verify R542000201
                Site.CaptureRequirementIfIsNotNull(
                    list.Flags2,
                    542000201,
                    @"Implementation does return this attribute[ListDefinitionCT.Flags2]. [In Appendix B: Product Behavior] (SharePoint Server 2016 and above support this behavior.)");
            }

            if (Common.IsRequirementEnabled(542000901, this.Site))
            {
                // Verify R542000901
                Site.CaptureRequirementIfIsNotNull(
                    list.RootFolderId,
                    542000901,
                    @"Implementation does return this attribute[ListDefinitionCT.RootFolderId]. [In Appendix B: Product Behavior] (SharePoint Server 2016 and above support this behavior.)");
            }

            if (Common.IsRequirementEnabled(542001001, this.Site))
            {
                // Verify R542001001
                Site.CaptureRequirementIfIsNotNull(
                    list.IrmSyncable,
                    542001001,
                    @"Implementation does return this attribute[ListDefinitionCT.IrmSyncable]. [In Appendix B: Product Behavior] (SharePoint Server 2016 and above support this behavior.)");
            }
        }

        /// <summary>
        /// Verify the requirements of the complex type ListDefinitionSchema.
        /// </summary>
        /// <param name="listDefinitionSchema">The actual ListDefinitionSchema.</param>
        private void VerifyListDefinitionSchema(ListDefinitionSchema listDefinitionSchema)
        {
            this.Site.Assert.IsNotNull(listDefinitionSchema, "The ListDefinitionSchema should not be null.");

            // Verify R1418
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1418,
                @"[ListDefinitionSchema]Specifies the results from retrieving a specified list.");

            if (listDefinitionSchema.Validation != null)
            {
                if (listDefinitionSchema.Validation.Message != null)
                {
                    bool isLengthLessThan1024 = listDefinitionSchema.Validation.Message.Length <= 1024;

                    // Verify MS-LISTSWS requirement: MS-LISTSWS_R2381 and MS-LISTSWS_R2380
                    // If the length of the Validation is not greater than 1024, then the following 
                    // requirement can be captured.              
                    if (Common.IsRequirementEnabled(2381, this.Site))
                    {
                        Site.CaptureRequirementIfIsTrue(
                            isLengthLessThan1024,
                            2381,
                            @"[ListDefinitionSchema.Validation]Implementation does not return characters longer than 1024, if this attribute present.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
                    }

                    if (Common.IsRequirementEnabled(2380, this.Site))
                    {
                        Site.CaptureRequirementIfIsTrue(
                            isLengthLessThan1024,
                            2380,
                            @"[ListDefinitionSchema.Validation.Message]Implementation does not return characters longer than 1024, if this attribute present.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
                    }
                }
            }

            // Verify R1419
            // ListDefinitionSchema should have been verified by the ListsSoap.XmlValidater method. 
            // Thus R1419 can be captured directly here.               
            Site.CaptureRequirement(
                1419,
                @"[The schema of ListDefinitionSchema defined as:]"
                + @"<s:complexType name=""ListDefinitionSchema"">"
                + @"  <s:complexContent>"
                + @"          <s:extension base=""tns:ListDefinitionCT"">"
                + @"            <s:sequence>"
                + @"              <s:element name=""Validation"" minOccurs=""0"">"
                + @"                <s:complexType>"
                + @"                  <s:attribute name=""Message"" type=""s:string"" use=""optional"" />"
                + @"                </s:complexType>"
                + @"              </s:element>"
                + @"              <s:element name=""ValidationDisplayNames"" minOccurs=""0"" type=""s:string"" />"
                + @"              <s:element name=""Fields"">"
                + @"                <s:complexType mixed=""true"">"
                + @"                  <s:sequence>"
                + @"                    <s:element name=""Field"" "
                + @"                               type=""core:FieldDefinition"" "
                + @"                               minOccurs=""0"" maxOccurs=""unbounded"" />"
                + @"                  </s:sequence>"
                + @"                </s:complexType>"
                + @"              </s:element>"
                + @"              <s:element name=""RegionalSettings"" >"
                + @"                <s:complexType mixed=""true"">"
                + @"                  <s:sequence>"
                + @"                    <s:element name=""Language"" type=""s:string"" />"
                + @"                    <s:element name=""Locale"" type=""s:string"" />"
                + @"                    <s:element name=""AdvanceHijri"" type=""s:string"" />"
                + @"                    <s:element name=""CalendarType"" type=""s:string"" />"
                + @"                    <s:element name=""Time24"" type=""s:string"" />"
                + @"                    <s:element name=""TimeZone"" type=""s:string"" />"
                + @"                    <s:element name=""SortOrder"" type=""s:string"" />"
                + @"                    <s:element name=""Presence"" type=""s:string"" />"
                + @"                  </s:sequence>"
                + @"                </s:complexType>"
                + @"              </s:element>"
                + @"              <s:element name=""ServerSettings"" >"
                + @"                <s:complexType mixed=""true"">"
                + @"                  <s:sequence>"
                + @"                    <s:element name=""ServerVersion"" type=""s:string"" />"
                + @"                    <s:element name=""RecycleBinEnabled"" type=""core:TRUEFALSE"" />"
                + @"                    <s:element name=""ServerRelativeUrl"" type=""s:string"" />"
                + @"                  </s:sequence>"
                + @"                </s:complexType>"
                + @"              </s:element>"
                + @"            </s:sequence>"
                + @"          </s:extension>"
                + @"  </s:complexContent>"
                + @"</s:complexType>");

            Site.Assert.IsNotNull(listDefinitionSchema, "The ListDefinitionSchema cannot be null");
            Site.Assert.IsNotNull(listDefinitionSchema.Fields, "The ListDefinitionSchema.Fields cannot be null");

            // Verify R1420
            if (listDefinitionSchema.Fields.Field != null)
            {
                if (listDefinitionSchema.Fields.Field.Length > 0)
                {
                    // If the response contains the "Field" element, its definition should have been verified by the ListsSoap.XmlValidater method. 
                    // Thus R1420 can be captured directly here.
                    Site.CaptureRequirement(
                        1420,
                        @"[ListDefinitionSchema.FieldDefinition: ]As specified in [MS-WSSFO2] section "
                        + "2.2.8.3.3.[A field definition describes the structure and format of a field that "
                        + "is used within a list or content type.]");
                }
            }

            // Verify R1421
            System.Globalization.CultureInfo[] cultures = System.Globalization.CultureInfo.GetCultures(System.Globalization.CultureTypes.AllCultures & ~System.Globalization.CultureTypes.NeutralCultures);

            bool isValidLCIDForLanguage = false;
            if (listDefinitionSchema.RegionalSettings.Language != null)
            {
                // If the response contains the "Language" element, and the value is a valid LCID, then the following requirement can be captured.
                foreach (System.Globalization.CultureInfo culture in cultures)
                {
                    if (listDefinitionSchema.RegionalSettings.Language == culture.LCID.ToString())
                    {
                        isValidLCIDForLanguage = true;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isValidLCIDForLanguage,
                    1421,
                    @"[ListDefinitionSchema.Language:] A valid language code identifier (LCID) as defined in [MS-LCID].");
            }

            // Verify R1422
            bool isValidLCIDForLocal = false;
            if (listDefinitionSchema.RegionalSettings.Locale != null)
            {
                // If the response contains the "Locale" element, and the value is a valid LCID, then the following requirement can be captured.
                foreach (System.Globalization.CultureInfo culture in cultures)
                {
                    if (listDefinitionSchema.RegionalSettings.Locale == culture.LCID.ToString())
                    {
                        isValidLCIDForLocal = true;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isValidLCIDForLocal,
                    1422,
                    @"[ListDefinitionSchema.Locale:] A valid language code identifier (LCID) as defined in [MS-LCID].");
            }

            // Verify R165
            // If the value is between -2 and 2, then the following requirement can be captured.
            int advanceHijri = int.Parse(listDefinitionSchema.RegionalSettings.AdvanceHijri);
            bool isVerifyR165 = advanceHijri >= -2 && advanceHijri <= 2;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR165,
                165,
                @"[ListDefinitionSchema.AdvanceHijri] An integer between -2 and 2.");

            // Verify R1423, R166
            if (!string.IsNullOrEmpty(listDefinitionSchema.RegionalSettings.CalendarType))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1423, R166.
                // The expected domain is defined in MS-WSSFO2 section 2.2.4.3. If the 
                // actual CalendarType value is contained in the expected domain of values, then the 
                // requirement can be captured.
                // And these are standard values which are not configurable
                string[] calendarTypeDomain = { "1", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "14", "15", "16" };
                List<string> listCalendarType = new List<string>(calendarTypeDomain);
                bool isVerifyR1423 = listCalendarType.Contains(listDefinitionSchema.RegionalSettings.CalendarType);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1423,
                    1423,
                    @"[ListDefinitionSchema.CalendarType:] Specifies the type of calendar.["
                    + "The only valid values of the Calendar Type are specified as follows."
                    + "Value Description"
                    + "1     Gregorian (localized)"
                    + "3     Japanese Emperor Era"
                    + "4     Taiwan Calendar"
                    + "5     Korean Tangun Era"
                    + "6     Hijri (Arabic Lunar)"
                    + "7     Thai"
                    + "8     Hebrew (Lunar)"
                    + "9     Gregorian (Middle East French)"
                    + "10    Gregorian (Arabic)"
                    + "11    Gregorian (Transliterated English)"
                    + "12    Gregorian (Transliterated French)"
                    + "14    Korean and Japan Lunar"
                    + "15    Chinese Lunar"
                    + "16    Saka Era]");
                bool isVerifyR166 = isVerifyR1423;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR166,
                    166,
                    @"[ListDefinitionSchema.CalendarType] See [MS-WSSFO2] section 2.2.3.3 for the "
                    + "different types of calendar supported.["
                    + "The only valid values of the Calendar Type are specified as follows."
                    + "Value Description"
                    + "1     Gregorian (localized)"
                    + "3     Japanese Emperor Era"
                    + "4     Taiwan Calendar"
                    + "5     Korean Tangun Era"
                    + "6     Hijri (Arabic Lunar)"
                    + "7     Thai"
                    + "8     Hebrew (Lunar)"
                    + "9     Gregorian (Middle East French)"
                    + "10    Gregorian (Arabic)"
                    + "11    Gregorian (Transliterated English)"
                    + "12    Gregorian (Transliterated French)"
                    + "14    Korean and Japan Lunar"
                    + "15    Chinese Lunar"
                    + "16    Saka Era]");
            }

            if (!string.IsNullOrEmpty(listDefinitionSchema.ServerSettings.RecycleBinEnabled))
            {
                // Verify R1209
                // If all the above requirements are verified, then the requirement can be 
                // captured.
                Site.CaptureRequirement(
                    1209,
                    @"[In Complex Types]The Complex type ListDefinitionSchema is used Specifies the "
                    + "results from retrieving a specified list.");

                if (Common.IsRequirementEnabled(2418, this.Site))
                {
                    // Verify R2418
                    Site.CaptureRequirementIfIsNull(
                        listDefinitionSchema.Validation,
                        2418,
                        @"Implementation does not return this attribute[ListDefinitionSchema.Validation]. [In Appendix B: Product Behavior] <26> Section 2.2.4.12: This attribute[ListDefinitionSchema.Validation] is not returned in Windows SharePoint Services 3.0.");
                }

                if (Common.IsRequirementEnabled(2420, this.Site))
                {
                    // Verify R2420
                    Site.CaptureRequirementIfIsNull(
                        listDefinitionSchema.Validation,
                        2420,
                        @"Implementation does not return this attribute[ListDefinitionSchema.Validation.Message]. [In Appendix B: Product Behavior] <27> Section 2.2.4.12: This attribute[ListDefinitionSchema.Validation.Message] is not returned in Windows SharePoint Services 3.0.");
                }

                if (Common.IsRequirementEnabled(2422, this.Site))
                {
                    // Verify R2422
                    Site.CaptureRequirementIfIsNull(
                        listDefinitionSchema.ValidationDisplayNames,
                        2422,
                        @"Implementation does not return this attribute[ListDefinitionSchema.ValidationDisplayNames]. [In Appendix B: Product Behavior] <28> Section 2.2.4.12: This attribute[ListDefinitionSchema.ValidationDisplayNames] is not returned in Windows SharePoint Services 3.0.");
                }
            }
        }

        /// <summary>
        /// A method used to verify "ServerChangeUnit" attribute does not return in GetListItemChangesSinceToken operation.
        /// </summary>
        /// <param name="change">A parameter represents the instance of ListItemChangeDefinition which will be verify whether contain "ServerChangeUnit" attribute</param>
        private void VerifyServerChangeUnitAttributeNotReturn(ListItemChangeDefinition change)
        {
            // If the ServerChangeUnit is null or empty in response of GetListItemChangesSinceToken operation.
            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(change.ServerChangeUnit),
                1444,
                @"[ServerChangeUnit:]This attribute is not returned when the element is "
                + "contained in the GetListItemChangesSinceTokenResult element.");
        }

        /// <summary>
        /// Verify the requirements of the complex type ListItemChangeDefinition.
        /// </summary>
        /// <param name="change">The actual ListItemChangeDefinition complex type.</param>
        private void VerifyListItemChangeDefinition(
            ListItemChangeDefinition change)
        {
            // Verify R1433
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1433,
                @"[The schema of ListItemChangeDefinition is defined as:]"
                + @"  <s:complexType name=""ListItemChangeDefinition"" mixed=""true"">"
                + @"  <s:attribute name=""ChangeType"" type=""tns:ChangeTypeEnum"" />"
                + @"  <s:attribute name=""AfterListId"" type=""core:UniqueIdentifierWithOrWithoutBraces"" />"
                + @"  <s:attribute name=""AfterItemId"" type=""s:unsignedInt"" />"
                + @"  <s:attribute name=""UniqueId"" type=""core:UniqueIdentifierWithOrWithoutBraces"" />"
                + @"  <s:attribute name=""MetaInfo_vti_clientid"" type=""s:string"" />"
                + @"  <s:attribute name=""ServerChangeUnit"" type=""s:string"" />"
                + @"</s:complexType>");

            // Verify the requirements of the ChangeTypeEnum simple type.
            this.VerifyChangeTypeEnum();

            Site.Assert.IsNotNull(change, "The ListItemChangeDefinition cannot be null");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1437
            // If the AfterListId is not null only when the ChangeType is MoveAway,             
            // then the following requirement can be captured.
            bool isVerifyR1437 = (change.ChangeType == ChangeTypeEnum.MoveAway) ^ (change.AfterListId == null);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1437,
                1437,
                @"[AfterListId]This MUST be set only for a change type of ChangeTypeEnum.MoveAway.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1439
            // If the AfterItemId is not 0 only when the ChangeType is MoveAway,             
            // then the following requirement can be captured.
            bool isVerifyR1439 = (change.ChangeType == ChangeTypeEnum.MoveAway) ^ (change.AfterItemId == 0);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1439,
                1439,
                @"[AfterItemId]This MUST be set only for a change type of ChangeTypeEnum.MoveAway.");

            if (Common.IsRequirementEnabled(2423, this.Site))
            {
                // Verify R2423
                Site.CaptureRequirementIfIsNull(
                    change.UniqueId,
                    2423,
                    @"Implementation does not return this attribute[ListItemChangeDefinition.UniqueId]. [In Appendix B: Product Behavior] <29> Section 2.2.4.13: This attribute is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2424, this.Site))
            {
                // Verify R2424
                Site.CaptureRequirementIfIsNull(
                    change.MetaInfo_vti_clientid,
                    2424,
                    @"Implementation does not return this attribute[ListItemChangeDefinition.MetaInfo_vti_clientid]. [In Appendix B: Product Behavior] <30> Section 2.2.4.13: This attribute is not returned in Windows SharePoint Services 3.0.");
            }

            if (Common.IsRequirementEnabled(2425, this.Site))
            {
                // Verify R2425
                Site.CaptureRequirementIfIsNull(
                    change.ServerChangeUnit,
                    2425,
                    @"[In Appendix B: Product Behavior] Implemementation does not return this attribute[ServerChangeUnit]. "
                    + "<31> Section 2.2.4.13: This attribute[ServerChangeUnit] is not returned in Windows SharePoint Services 3.0.");
            }
        }

        /// <summary>
        /// Verify the requirements of the complex type UpdateListFieldResults.
        /// </summary>
        /// <param name="results">The actual UpdateListFieldResults complex type.</param>
        private void VerifyUpdateListFieldResults(UpdateListResponseUpdateListResultResults results)
        {
            // Verify R1445
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1445,
                @"[UpdateListFieldResults]Specifies the results from  an Add, Update, or Delete "
                + "operation on a list's fields.");

            // Verify R1446
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured   
            Site.CaptureRequirement(
                1446,
                @"[The schema of  UpdateListFieldResults is defined as :]"
                + @"<s:complexType name=""UpdateListFieldResults"">"
                + @"  <s:sequence>"
                + @"    <s:element name=""Method"" minOccurs=""0"" maxOccurs=""unbounded"">"
                + @"      <s:complexType mixed=""true"">"
                + @"        <s:sequence>"
                + @"          <s:element name=""ErrorCode"" type=""s:string"" />"
                + @"          <s:element name=""ErrorText"" type=""s:string"" minOccurs=""0"" />"
                + @"          <s:element name=""Field"" type=""core:FieldDefinition"" minOccurs=""0""/>"
                + @"        </s:sequence>"
                + @"        <s:attribute name=""ID"" type=""s:string"" />"
                + @"      </s:complexType>"
                + @"     </s:element>"
                + @"  </s:sequence>"
                + @"</s:complexType>");

            Site.Assert.IsNotNull(results, "The UpdateListFieldResults cannot be null");

            // Verify R1448
            string strErrorCode = string.Empty;
            if ((results.NewFields != null) && (results.NewFields.Length > 0))
            {
                strErrorCode = results.NewFields[0].ErrorCode;
            }
            else if ((results.UpdateFields != null) && (results.UpdateFields.Length > 0))
            {
                strErrorCode = results.UpdateFields[0].ErrorCode;
            }
            else if ((results.DeleteFields != null) && (results.DeleteFields.Length > 0))
            {
                strErrorCode = results.DeleteFields[0].ErrorCode;
            }

            if (!string.IsNullOrEmpty(strErrorCode))
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1448
                // If the ErrorCode can be parsed to int, then the following requirement can be captured.
                int errorCode = int.MinValue;
                bool isVerifyR1448 = int.TryParse(
                    strErrorCode.Substring(2),
                    System.Globalization.NumberStyles.HexNumber,
                    null,
                    out errorCode);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1448,
                    1448,
                    @"[ErrorCode:] The string representation of a hexadecimal number indicating "
                    + "whether the operation succeeded or failed.");
            }

            // Verify R1451
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1451,
                @"[Field: ]As specified in [MS-WSSFO2] section 2.2.8.3.3.[A field definition describes "
                + "the structure and format of a field that is used within a list or content type.]");

            // Verify R1211
            // If all the above requirements are verified, then the requirement can be 
            // captured.
            Site.CaptureRequirement(
                1211,
                @"[In Complex Types]The Complex type UpdateListFieldResults is used Specifies the "
                + "results from an Add, Update, or Delete operation on a list's fields.");
        }

        #endregion

        #region Capture Adapter requirements of Simple Types

        /// <summary>
        /// Verify the requirements of the simple type ChangeTypeEnum.
        /// </summary>
        private void VerifyChangeTypeEnum()
        {
            // Verify R1466
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1466,
                @"[ChangeTypeEnum]Specifies the type of changes returned when a protocol "
                + "client requests changes to list items.");

            // Verify R1473
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1473,
                @"[ChangeTypeEnum]Specifies the type of changes returned when a protocol "
                + "client requests changes to list items.");

            // Verify R1474
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1474,
                @"[The schema of ChangeTypeEnum is defined as:]"
                + @"<s:simpleType name=""ChangeTypeEnum"">"
                + @"  <s:restriction base=""s:string"">"
                + @"    <s:enumeration value=""Delete"" />"
                + @"    <s:enumeration value=""InvalidToken"" />"
                + @"    <s:enumeration value=""Restore"" />"
                + @"    <s:enumeration value=""MoveAway"" />"
                + @"    <s:enumeration value=""SystemUpdate"" />"
                + @"    <s:enumeration value=""Rename"" />"
                + @"  </s:restriction>"
                + @"</s:simpleType>");
        }

        /// <summary>
        /// Verify the requirements of the simple type EnumViewAttributes.
        /// </summary>
        private void VerifyEnumViewAttributes()
        {
            // Verify R274
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                274,
                @"[In EnumViewAttributes] The values [Recursive, RecursiveAll, FilesOnly] here MUST be used by the protocol server to restrict the data returned in document libraries.");
        }

        /// <summary>
        /// Verify the requirements of the simple type TRUEONLY.
        /// </summary>
        private void VerifyTRUEONLY()
        {
            // Verify R1472
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1472,
                @"[TRUEONLY]Specifies that a particular attribute is restricted to only the value "
                + @"""TRUE"".");

            // Verify R1508
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1508,
                @"[TRUEONLY]Specifies that a particular attribute is restricted to only the value "
                + @"""TRUE"".");

            // Verify R1509
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1509,
                @"[The schema of TRUEONLY is defined as:]<s:simpleType name=""TRUEONLY"">"
                + @"  <s:restriction base=""s:string"">"
                + @"    <s:enumeration value=""TRUE"" />"
                + @"  </s:restriction>"
                + @"</s:simpleType>");
        }

        #endregion

        #region Capture Adapter requirements of Operations

        /// <summary>
        /// Verify the message syntax of AddAttachment operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="attachmentRelativeUrl">The returned SOAP result.</param>
        private void VerifyAddAttachmentOperation(string attachmentRelativeUrl)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                attachmentRelativeUrl,
                "The result of AddAttachment operation must not be null.");

            // Verify R1549
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1549,
                @"[The schema of Addattachment is defined as:]"
                + @"<wsdl:operation name=""AddAttachment"">"
                + @"    <wsdl:input message=""AddAttachmentSoapIn"" />"
                + @"    <wsdl:output message=""AddAttachmentSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R288
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                288,
                @"[AddAttachment]The protocol client sends an AddAttachmentSoapIn request "
                + "message and the server responds with an AddAttachmentSoapOut response "
                + "message.");

            // Verify R1561
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1561,
                @"[AddAttachmentSoapOut]The SOAP Body contains an AddAttachmentResponse "
                + "element.");

            // Verify R1569
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1569,
                @"[AddAttachmentResponse]<s:element name=""AddAttachmentResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""AddAttachmentResult"" type=""s:string"" minOccurs=""0""/>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");
        }

        /// <summary>
        /// Verify the message syntax of AddDiscussionBoardItem operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="addDiscussionBoardItemResult">The returned SOAP result.</param>
        private void VerifyAddDiscussionBoardItemOperation(AddDiscussionBoardItemResponseAddDiscussionBoardItemResult addDiscussionBoardItemResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                addDiscussionBoardItemResult,
                "The result of AddAttachment operation must not be null.");

            // Verify R1571
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1571,
                @"[The schema of AddDiscussionBoardItem is defined as:]"
                + @"<wsdl:operation name=""AddDiscussionBoardItem"">"
                + @"    <wsdl:input message=""AddDiscussionBoardItemSoapIn"" />"
                + @"    <wsdl:output message=""AddDiscussionBoardItemSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R311
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                311,
                @"[In AddDiscussionBoardItem operation] [If the protocol client sends an "
                + "AddDiscussionBoardItemSoapIn request message] the protocol server "
                + "responds with an AddDiscussionBoardItemSoapOut response message.");

            // Verify R1577
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1577,
                @"[AddDiscussionBoardItemSoapOut]The SOAP Body contains an "
                + "AddDiscussionBoardItemResponse element.");

            // Verify R1581
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1581,
                @"[The schema of AddDiscussionBoardItemResponse is defined as:]"
                + @"<s:element name=""AddDiscussionBoardItemResponse"">  "
                + @"  <s:complexType>    "
                + @"    <s:sequence>      "
                + @"      <s:element name=""AddDiscussionBoardItemResult"" minOccurs=""0""> "
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""listitems"" >"
                + @"              <s:complexType>"
                + @"                <s:sequence>"
                + @"                  <s:any />"
                + @"                </s:sequence>"
                + @"                <s:anyAttribute />"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R330
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                330,
                @"[In AddDiscussionBoardItem operation] [In AddDiscussionBoardItemResponse element] [In listitems field] "
                + "The protocol server response included in the listitems element is modeled on a persistence format as specified "
                + "in [MS-PRSTFR], excluding the <s:schema> element.");

            // Verify R334
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                334,
                @"[In AddDiscussionBoardItem operation] [In AddDiscussionBoardItemResponse element] [In listitems field] "
                + "The listitems element includes attributes describing the namespaces for the persistence format");

            // Verify R1584
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1584,
                @"[AddDiscussionBoardItemResponse]listitems contains an inner element named "
                + "rs:data, which is of type DataDefinition.");

            // Verify the requirements of the DataDefinition complex type.
            if (addDiscussionBoardItemResult.listitems.data != null)
            {
                this.VerifyDataDefinition(addDiscussionBoardItemResult.listitems.data);
            }
        }

        /// <summary>
        /// Verify the message syntax of AddList operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="addListResult">The returned SOAP result.</param>
        private void VerifyAddListOperation(AddListResponseAddListResult addListResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(addListResult, "The result of AddList operation must not be null.");
            Site.Assume.IsNotNull(addListResult.List, "AddListResponseAddListResult.List must not be null.");

            // Verify R1585
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1585,
                @"[The schema of AddList is defined as:]<wsdl:operation name=""AddList"">"
                + @"    <wsdl:input message=""AddListSoapIn"" />"
                + @"    <wsdl:output message=""AddListSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R340
            // The response have been received successfully, then the following 
            // requirements can be captured. If any of the following requirements is fail, 
            // the response can't be received successfully.
            Site.CaptureRequirement(
                340,
                @"[In AddList operation] [If the protocol client sends an AddListSoapIn request "
                + "message] the server responds with an AddListSoapOut response message.");

            // Verify R1593
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1593,
                @"[AddListSoapOut]The SOAP Body contains an AddListResponse element.");

            // Verify R1598
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1598,
                @"[The schema of AddListResponse is defined as:]"
                + @"<s:element name=""AddListResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element minOccurs=""0"" maxOccurs=""1"" name=""AddListResult"" >"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""List"" type=""tns:ListDefinitionSchema"" />"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify the requirements of the ListDefinitionSchema and ListDefinitionCT complex type.
            if (addListResult.List != null)
            {
                this.VerifyListDefinitionCT(addListResult.List);
                this.VerifyListDefinitionSchema(addListResult.List);
            }
        }

        /// <summary>
        /// Verify the message syntax of AddListFromFeature operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="addListFromFeatureResult">The result of the operation.</param>
        private void VerifyAddListFromFeatureOperation(
            AddListFromFeatureResponseAddListFromFeatureResult addListFromFeatureResult)
        {
            // Verify R1600
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1600,
                @"[The schema of AddListFromFeature is defined as:]"
                + @"<wsdl:operation name=""AddListFromFeature"">"
                + @"    <wsdl:input message=""AddListFromFeatureSoapIn"" />"
                + @"    <wsdl:output message=""AddListFromFeatureSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R356
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                356,
                @"[In AddListFromFeature operation] [If the protocol client sends an "
                + "AddListFromFeatureSoapIn request message,] the server responds with an "
                + "AddListFromFeatureSoapOut response message.");

            // Verify R1613
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured   
            Site.CaptureRequirement(
                1613,
                @"[AddListFromFeatureSoapOut]The SOAP Body contains an "
                + "AddListFromFeatureResponse element");

            // Verify R1621
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1621,
                @"[The schema of AddListFromFeatureResponse is defined as:] "
                + @"<s:element minOccurs=""0"" maxOccurs=""1"" name=""AddListFromFeatureResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""AddListFromFeatureResult"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""List"" type=""tns:ListDefinitionSchema""  />"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify the requirements of the ListDefinitionCT complex type.
            this.VerifyListDefinitionCT(addListFromFeatureResult.List);

            // Verify the requirements of the ListDefinitionSchema complex type.
            if (addListFromFeatureResult.List != null)
            {
                // Verify R1622
                Site.CaptureRequirement(
                    1622,
                    @"[AddListFromFeatureResponse]AddListFromFeatureResult: Contains information about the properties and schema of the list created by the AddListFromFeature operation. See section 2.2.4.12 for more details.");

                this.VerifyListDefinitionSchema(addListFromFeatureResult.List);
            }
        }

        /// <summary>
        /// Verify the message syntax of ApplyContentTypeToList operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="applyContentTypeToListResult">The result of the operation.</param>
        private void VerifyApplyContentTypeToListOperation(
            ApplyContentTypeToListResponseApplyContentTypeToListResult applyContentTypeToListResult)
        {
            // Verify R1624
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1624,
                @"[The schema of ApplyContentTypeToList is defined as:]"
                + @"<wsdl:operation name=""ApplyContentTypeToList"">"
                + @"    <wsdl:input message=""ApplyContentTypeToListSoapIn"" />"
                + @"    <wsdl:output message=""ApplyContentTypeToListSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R372
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                372,
                @"[In ApplyContentTypeToList operation] [If the protocol client sends an "
                + "ApplyContentTypeToListSoapIn request message] the protocol server "
                + "responds with an ApplyContentTypeToListSoapOut response message.");

            // Verify R1629
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1629,
                @"[ApplyContentTypeToListSoapOut]The SOAP Body contains an "
                + "ApplyContentTypeToListResponse element.");

            // Verify R1635
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1635,
                @"[The schema of ApplyContentTypeToListResponse is defined as:] "
                + @"<s:element minOccurs=""0"" maxOccurs=""1"" name=""ApplyContentTypeToListResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""ApplyContentTypeToListResult"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Success"">"
                + @"              <s:complexType/>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R384
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                applyContentTypeToListResult,
                "The result of ApplyContentTypeToList operation must not be null.");

            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                applyContentTypeToListResult,
                384,
                @"[In ApplyContentTypeToList operation] If the operation succeeds, an "
                + "ApplyContentTypeToListResult MUST be returned.");
        }

        /// <summary>
        /// Verify the message syntax of CheckInFile operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="result">The result of the operation.</param>
        private void VerifyCheckInFileOperation(bool result)
        {
            if (result)
            {
                // Verify R1639
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1639,
                    @"[The schema of CheckInFile is defined as:]<wsdl:operation name=""CheckInFile"">"
                    + @"<wsdl:input message=""CheckInFileSoapIn"" />"
                    + @"<wsdl:output message=""CheckInFileSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R389
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    389,
                    @"[In CheckInFile operation] [If the protocol client sends a CheckInFileSoapIn "
                    + "request message] the protocol server responds with a CheckInFileSoapOut "
                    + "response message.");

                // Verify R1648
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1648,
                    @"[CheckInFileSoapOut]The SOAP Body contains a CheckInFileResponse element.");

                // Verify R1658
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1658,
                    @"[The schema of CheckInFileResponse is defined as:]"
                    + @"<s:element name=""CheckInFileResponse"">"
                    + @"  <s:complexType>"
                    + @"    <s:sequence>"
                    + @"      <s:element name=""CheckInFileResult"" type=""s:boolean""/>"
                    + @"    </s:sequence>"
                    + @"  </s:complexType>"
                    + @"</s:element>");

                // Verify R1659
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1659,
                    @"[CheckInFileResponse]CheckInFileResult: The value of the CheckInFileResult "
                    + "specifies whether the call is successful or not.");
            }
        }

        /// <summary>
        /// Verify the message syntax of CheckOutFile operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="result">The result of the operation.</param>
        private void VerifyCheckOutFileOperation(bool result)
        {
            if (result)
            {
                // Verify R1662
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1662,
                    @"[The schema of CheckOutFile is defined as:]"
                    + @"<wsdl:operation name=""CheckOutFile"">"
                    + @"    <wsdl:input message=""CheckOutFileSoapIn"" />"
                    + @"    <wsdl:output message=""CheckOutFileSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R405
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    405,
                    @"[In CheckOutFile operation] [If the protocol client sends a CheckOutFileSoapIn "
                    + "request message] the protocol server responds with a CheckOutFileSoapOut "
                    + "response message.");

                // Verify R1671
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1671,
                    @"[CheckOutFileSoapOut]The SOAP Body contains a CheckOutFileResponse "
                    + "element");

                // Verify R1677
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1677,
                    @"[The schema of CheckOutFileResponse is defined as:]"
                    + @"<s:element name=""CheckOutFileResponse"">"
                    + @"  <s:complexType>"
                    + @"    <s:sequence>"
                    + @"      <s:element name=""CheckOutFileResult"" type=""s:boolean""/>"
                    + @"    </s:sequence>"
                    + @"  </s:complexType>"
                    + @"</s:element>");

                // Verify R1678
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1678,
                    @"[CheckOutFileResponse]CheckOutFileResult: The value of CheckOutFileResult "
                    + "specifies whether the call is successful or not.");
            }
        }

        /// <summary>
        /// Verify the message syntax of CreateContentType operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="createContentTypeResult">The result of the operation.</param>
        private void VerifyCreateContentTypeOperation(string createContentTypeResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                createContentTypeResult,
                "The result of CreateContentType operation must not be null.");

            // Verify R1681
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured
            Site.CaptureRequirement(
                1681,
                @"[The schema of CreateContentType is defined as:]"
                + @"<wsdl:operation name=""CreateContentType"">"
                + @"    <wsdl:input message=""CreateContentTypeSoapIn"" />"
                + @"    <wsdl:output message=""CreateContentTypeSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R420
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                420,
                @"[In CreateContentType operation] [If the protocol client sends a "
                + "CreateContentTypeSoapIn request message] the protocol server responds "
                + "with a CreateContentTypeSoapOut response message.");

            // Verify R1688
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1688,
                @"[CreateContentTypeSoapOut]The SOAP body contains a "
                + "CreateContentTypeResponse element");

            // Verify R1695
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured  
            Site.CaptureRequirement(
                1695,
                @"[CreateContentTypeResponse] <s:element name=""CreateContentTypeResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element minOccurs=""0"" maxOccurs=""1"" name=""CreateContentTypeResult"" type=""s:string"" />"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R460
            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                createContentTypeResult,
                460,
                @"[In CreateContentType operation] [If no error condition, as specified in the "
                + "preceding section, causes the protocol server to return a SOAP fault] "
                + "CreateContentTypeResult MUST be returned.");
        }

        /// <summary>
        /// Verify the message syntax of DeleteAttachment operation.
        /// </summary>
        private void VerifyDeleteAttachmentOperation()
        {
            // Verify 1697
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1697,
                @"[The schema of DeleteAttachment is defined as:]"
                + @"<wsdl:operation name=""DeleteAttachment"">"
                + @"    <wsdl:input message=""DeleteAttachmentSoapIn"" />"
                + @"    <wsdl:output message=""DeleteAttachmentSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R465
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                465,
                @"[In DeleteAttachment operation] [If the protocol client sends a "
                + "DeleteAttachmentSoapIn request message] the protocol server responds with "
                + "a DeleteAttachmentSoapOut response message.");

            // Verify R1706
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1706,
                @"[DeleteAttachmentSoapOut]The SOAP Body contains a "
                + "DeleteAttachmentResponse element.");

            // Verify R1711
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1711,
                @"[The schema of DeleteAttachmentResponse is defined as:]"
                + @"<s:element name=""DeleteAttachmentResponse"">"
                + @"  <s:complexType/>"
                + @"</s:element>");
        }

        /// <summary>
        /// Verify the message syntax of DeleteContentType operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="deleteContentTypeResult">The result of the operation.</param>
        private void VerifyDeleteContentTypeOperation(DeleteContentTypeResponseDeleteContentTypeResult deleteContentTypeResult)
        {
            // Verify R1712
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1712,
                @"[The schema of DeleteContentType is defined as:]"
                + @"<wsdl:operation name=""DeleteContentType"">"
                + @"    <wsdl:input message=""DeleteContentTypeSoapIn"" />"
                + @"    <wsdl:output message=""DeleteContentTypeSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R484
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                484,
                @"[In DeleteContentType operation] [If the protocol client sends a "
                + "DeleteContentTypeSoapIn request message] the protocol server responds "
                + "with a DeleteContentTypeSoapOut response message.");

            // Verify R1718
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1718,
                @"[DeleteContentTypeSoapOut]The SOAP Body contains a "
                + "DeleteContentTypeResponse element.");

            // Verify R1722
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1722,
                @"[The schema of DeleteContentTypeResponse is defined as:]"
                + @"<s:element name=""DeleteContentTypeResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""DeleteContentTypeResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Success"">"
                + @"              <s:complexType />"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R495
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                deleteContentTypeResult,
                "The result of DeleteContentType operation must not be null.");

            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                deleteContentTypeResult,
                495,
                @"[In DeleteContentType operation] If the operation [DeleteContentType] "
                + "succeeds, a DeleteContentTypeResult MUST be returned.");
        }

        /// <summary>
        /// Verify the message syntax of DeleteContentTypeXmlDocument operation when the response 
        /// is received successfully.
        /// </summary>
        /// <param name="deleteContentTypeXmlDocumentResult">The result of the operation.</param>
        private void VerifyDeleteContentTypeXmlDocumentOperation(DeleteContentTypeXmlDocumentResponseDeleteContentTypeXmlDocumentResult deleteContentTypeXmlDocumentResult)
        {
            // Verify R1724
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1724,
                @"[The schema of DeleteContentTypeXmlDocument is defined as:]"
                + @"<wsdl:operation name=""DeleteContentTypeXmlDocument"">"
                + @"    <wsdl:input message=""DeleteContentTypeXmlDocumentSoapIn"" />"
                + @"    <wsdl:output message=""DeleteContentTypeXmlDocumentSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R501
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                501,
                @"[In DeleteContentTypeXmlDocument operation] [If the protocol client sends "
                + "a DeleteContentTypeXmlDocumentSoapIn request message] the protocol "
                + "server responds with a DeleteContentTypeXmlDocumentSoapOut response "
                + "message.");

            // Verify R1730
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1730,
                @"[DeleteContentTypeXmlDocumentSoapOut]The SOAP Body contains a "
                + "DeleteContentTypeXmlDocumentResponse element.");

            // Verify R1735
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1735,
                @"[DeleteContentTypeXmlDocumentResponse]"
                + @"<s:element name=""DeleteContentTypeXmlDocumentResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""DeleteContentTypeXmlDocumentResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Success"" minOccurs=""0"">"
                + @"              <s:complexType />"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R513
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                deleteContentTypeXmlDocumentResult,
                "The result of DeleteContentTypeXmlDocument operation must not be null.");

            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                deleteContentTypeXmlDocumentResult,
                513,
                @"[In DeleteContentTypeXmlDocument operation] If the operation succeeds, a "
                + "DeleteContentTypeXmlDocumentResult MUST be returned.");
        }

        /// <summary>
        /// Verify the message syntax of DeleteList operation when the response is received 
        /// successfully.
        /// </summary>
        private void VerifyDeleteListOperation()
        {
            // Verify R1737
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1737,
                @"[The schema of DeleteList is defined as:]<wsdl:operation name=""DeleteList"">"
                + @"    <wsdl:input message=""DeleteListSoapIn"" />"
                + @"    <wsdl:output message=""DeleteListSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R519
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                519,
                @"[In DeleteList operation] [If the protocol client sends a DeleteListSoapIn "
                + "request message] The protocol server MUST responds with a "
                + "DeleteListSoapOut response message.");

            // Verify R1742
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1742,
                @"[DeleteListSoapOut]The SOAP Body contains a DeleteListResponse element.");

            // Verify R528
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                528,
                @"[In DeleteList operation] This element [DeleteListResponse] contains a single "
                + "element that MUST be sent by the site if the DeleteList operation succeeds.");

            // Verify R1745
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1745,
                @"[The schema of DeleteListResponse is defined as:]"
                + @"<s:element name=""DeleteListResponse"">"
                + @"  <s:complexType/>"
                + @"</s:element>");
        }

        /// <summary>
        /// Verify the message syntax of GetAttachmentCollection operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="getAttachmentCollectionResult">The result of the operation.</param>
        private void VerifyGetAttachmentCollectionOperation(
            GetAttachmentCollectionResponseGetAttachmentCollectionResult getAttachmentCollectionResult)
        {
            this.Site.Assert.IsNotNull(getAttachmentCollectionResult, "Should get a correct response from GetAttachmentCollection operation.");

            // Verify R548
            this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: getAttachmentCollectionResult.Attachments.Length[{0}] for requirement #R548",
                    getAttachmentCollectionResult.Attachments.Length);

            this.Site.CaptureRequirementIfIsNotNull(
                       getAttachmentCollectionResult.Attachments,
                       548,
                       @"[In GetAttachmentCollection operation] If the operation succeeds, the protocol server MUST return the collection of attachment URLs for the specified list item.");

            // Verify R1746
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1746,
                @"[The schema of GetAttachmentCollection is defined as:]"
                + @"<wsdl:operation name=""GetAttachmentCollection"">"
                + @"    <wsdl:input message=""GetAttachmentCollectionSoapIn"" />"
                + @"    <wsdl:output message=""GetAttachmentCollectionSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R533
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                533,
                @"[In GetAttachmentCollection operation] [If the protocol client sends a "
                + "GetAttachmentCollectionSoapIn request message] the protocol server "
                + "responds with a GetAttachmentCollectionSoapOut response message.");

            // Verify R1753
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1753,
                @"[GetAttachmentCollectionSoapOut]The SOAP Body contains a "
                + "GetAttachmentCollectionResponse element.");

            // Verify R1758
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1758,
                @"[The schema of GetAttachmentCollectionResponse is defined as:]"
                + @"<s:element name=""GetAttachmentCollectionResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""GetAttachmentCollectionResult"" minOccurs=""0"">"
                + @"          <s:complexType mixed=""true"">"
                + @"            <s:sequence>"
                + @"              <s:element name=""Attachments"">"
                + @"                <s:complexType>"
                + @"                  <s:sequence>"
                + @"                    <s:element name=""Attachment"" type=""s:string"" minOccurs=""0""  "
                + @"                               maxOccurs=""unbounded"">"
                + @"                    </s:element>"
                + @"                  </s:sequence>"
                + @"                </s:complexType>"
                + @"              </s:element>"
                + @"            </s:sequence>"
                + @"          </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");
        }

        /// <summary>
        /// Verify the message syntax of GetList operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="getListResult">The returned SOAP result.</param>
        private void VerifyGetListOperation(GetListResponseGetListResult getListResult)
        {
            // Verify R1763
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1763,
                @"[The schema of GetList is defined as:]<wsdl:operation name=""GetList"">"
                + @"    <wsdl:input message=""GetListSoapIn"" />"
                + @"    <wsdl:output message=""GetListSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R554
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                554,
                @"[In GetList operation] [If the protocol client sends a GetListSoapIn request "
                + "message]The protocol server MUST respond with a GetListSoapOut response "
                + "message.");

            // Verify R1768
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1768,
                @"[GetListSoapOut]The SOAP Body contains a GetListResponse element.");

            // Verify R1771
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured
            Site.CaptureRequirement(
                1771,
                @"[GetListResponse]<s:element name=""GetListResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""GetListResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""List"" type=""tns:ListDefinitionSchema"" />"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            this.VerifyListDefinitionCT(getListResult.List);

            // Verify the requirements of the ListDefinitionSchema complex type.
            if (getListResult.List != null)
            {
                this.VerifyListDefinitionSchema(getListResult.List);
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListAndView operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="getListAndViewResult">The returned SOAP result.</param>
        private void VerifyGetListAndViewOperation(GetListAndViewResponseGetListAndViewResult getListAndViewResult)
        {
            Site.Assert.IsNotNull(
                getListAndViewResult,
                "The return result of this operation cannot be NULL");

            // Verify R1775
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1775,
                @"[The schema of GetListAndView is defined as:]"
                + @"<wsdl:operation name=""GetListAndView"">"
                + @"    <wsdl:input message=""GetListAndViewSoapIn"" />"
                + @"    <wsdl:output message=""GetListAndViewSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R567
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                567,
                @"[In GetListAndView operation] [If the protocol client sends a "
                + "GetListAndViewSoapIn request message] The server MUST respond with a "
                + "GetListAndViewSoapOut response message.");

            // Verify R1783
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1783,
                @"The SOAP Body contains a GetListAndViewResponse element.");

            // Verify R1787
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1787,
                @"[The schema of GetListAndViewResponse is defined as:]<s:element name=""GetListAndViewResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""GetListAndViewResult"" minOccurs=""0"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""ListAndView"">
                              <s:complexType mixed=""true"">
                                <s:sequence>
                                  <s:element name=""List"" type=""tns:ListDefinitionSchema"" />
                                  <s:element name=""View"" type=""core:ViewDefinition"" />
                                </s:sequence>
                              </s:complexType>
                            </s:element>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>
                  </s:complexType>
                </s:element>");

            // Verify R1789
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1789,
                @"[GetListAndViewResponse]ListDefinitionSchema: Specifies the schema for the list");

            // Verify R1790
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1790,
                @"[GetListAndViewResponse]ViewDefinition: Specifies the schema for the view.");

            Site.Assert.IsNotNull(getListAndViewResult, "The return result of this operation cannot be NULL");

            this.VerifyListDefinitionCT(getListAndViewResult.ListAndView.List);

            // Verify the requirements of the ListDefinitionSchema complex type.
            if (getListAndViewResult.ListAndView.List != null)
            {
                this.VerifyListDefinitionSchema(getListAndViewResult.ListAndView.List);
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListCollection operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="getListCollectionResult">The returned SOAP result.</param>
        private void VerifyGetListCollectionOperation(GetListCollectionResponseGetListCollectionResult getListCollectionResult)
        {
            Site.Assert.IsNotNull(getListCollectionResult, "The return value of this operation should have instance");

            // Verify R1791
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1791,
                @"[GetListCollection]<wsdl:operation name=""GetListCollection"">"
                + @"    <wsdl:input message=""GetListCollectionSoapIn"" />"
                + @"    <wsdl:output message=""GetListCollectionSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R581
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                581,
                @"[In GetListCollection operation] [If the protocol client sends a "
                + "GetListCollectionSoapIn request message] the server responds with a "
                + "GetListCollectionSoapOut response message.");

            // Verify R1795
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1795,
                @"[GetListCollectionSoapOut]The SOAP Body contains a GetListCollectionResponse "
                + "element.");

            // Verify R1797
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1797,
                @"[The schema of GetListCollectionResponse is defined as:]"
                + @"<s:element name=""GetListCollectionResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""GetListCollectionResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Lists"">"
                + @"              <s:complexType>"
                + @"                <s:sequence>"
                + @"                  <s:element name=""List"" type=""tns:ListDefinitionCT"" minOccurs=""0"" "
                + @"                             maxOccurs=""unbounded""/>"
                + @"                </s:sequence>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify the requirements of the ListDefinitionCT complex type.
            foreach (ListDefinitionCT listDefinitionCTItem in getListCollectionResult.Lists)
            {
                this.VerifyListDefinitionCT(listDefinitionCTItem);

                if (listDefinitionCTItem.HasRelatedLists != null)
                {
                    this.Site.Assert.AreEqual<string>(string.Empty, listDefinitionCTItem.HasRelatedLists, "[ListDefinitionCT.HasRelatedLists] When it is returned in GetListCollection (section 3.1.4.17) this value will be an empty string.");
                    //Verify MS-LISTSWS requirement: MS-LISTSWS_R3010002
                    Site.CaptureRequirement(
                        3010002,
                        @"[ListDefinitionCT.HasRelatedLists] When it is returned by GetListCollection (section 3.1.4.17) this value will be an empty string.");
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListContentType operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="getListContentTypeResult">The result of the operation.</param>
        private void VerifyGetListContentTypeOperation(GetListContentTypeResponseGetListContentTypeResult getListContentTypeResult)
        {
            // Verify R1803
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1803,
                @"[The schema of GetListContentType is defined as:]"
                + @"<wsdl:operation name=""GetListContentType"">"
                + @"    <wsdl:input message=""GetListContentTypeSoapIn"" />"
                + @"    <wsdl:output message=""GetListContentTypeSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R590
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                590,
                @"[In GetListContentType operation] [If the protocol client sends a "
                + "GetListContentTypeSoapIn request message] the protocol server responds "
                + "with a GetListContentTypeSoapOut response message.");

            // Verify R1809
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1809,
                @"[GetListContentTypeSoapOut]The SOAP Body contains a "
                + "GetListContentTypeResponse element.");

            // Verify R1813
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1813,
                @"[GetListContentTypeResponse] "
                + @"<s:element name=""GetListContentTypeResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"       <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetListContentTypeResult"">"
                + @"        <s:complexType mixed=""true"">"
                + @"        <s:complexType >"
                + @"            <s:sequence>"
                + @"              <s:element name=""ContentType"" type=""core:ContentTypeDefinition""/>"
                + @"            </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1814
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(getListContentTypeResult, "The result of GetListContentType operation must not be null.");

            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                getListContentTypeResult,
                1814,
                @"[GetListContentTypeResponse]GetListContentTypeResult: The container for the "
                + "returned content type data.");
        }

        /// <summary>
        /// Verify the message syntax of GetListContentTypes operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="getListContentTypesResult">The result of the operation.</param>
        private void VerifyGetListContentTypesOperation(GetListContentTypesResponseGetListContentTypesResult getListContentTypesResult)
        {
            if (getListContentTypesResult != null)
            {
                // Verify R1817
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1817,
                    @"[The schema of GetListContentTypes is defined as:]"
                    + @"<wsdl:operation name=""GetListContentTypes"">"
                    + @"    <wsdl:input message=""GetListContentTypesSoapIn"" />"
                    + @"    <wsdl:output message=""GetListContentTypesSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R605
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    605,
                    @"[In GetListContentTypes operation] [If the protocol client sends a "
                    + "GetListContentTypesSoapIn message] the protocol server responds with a "
                    + "GetListContentTypesSoapOut message.");

                // Verify R994
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                this.Site.CaptureRequirement(
                                    "MS-WSSCAML",
                                    994,
                                    @"[The schema definition of XmlDocumentDefinitionCollection is as follows:]<xs:complexType name=""XmlDocumentDefinitionCollection"">
                                        <xs:sequence>
                                          <xs:element name=""XmlDocument"" type=""XmlDocumentDefinition"" minOccurs=""0"" maxOccurs=""unbounded"" />
                                        </xs:sequence>
                                      </xs:complexType>");

                // Verify R1821
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1821,
                    @"[GetListContentTypesSoapOut]The SOAP Body contains a "
                    + @"GetListContentTypesResponse element.");

                // Verify R1829
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1829,
                    @"[The schema of GetListContentTypesResponse is defined as:]"
                    + @"<s:element name=""GetListContentTypesResponse"">"
                    + @"  <s:complexType>"
                    + @"    <s:sequence>"
                    + @"      <s:element name=""GetListContentTypesResult"" minOccurs=""0"">"
                    + @"        <s:complexType mixed=""true"">"
                    + @"          <s:sequence>"
                    + @"            <s:element name=""ContentTypes"" >"
                    + @"              <s:complexType>"
                    + @"                <s:sequence>"
                    + @"                  <s:element name=""ContentType"" maxOccurs=""unbounded"">"
                    + @"                    <s:complexType>"
                    + @"                      <s:sequence>"
                    + @"                        <s:element name=""XmlDocuments"" "
                    + @"                                   type=""core:XmlDocumentDefinitionCollection"" "
                    + @"                                   minOccurs=""0"">"
                    + @"                        </s:element>"
                    + @"                      </s:sequence>"
                    + @"                      <s:attribute name=""Name"" type=""s:string"" use=""required"" />"
                    + @"                      <s:attribute name=""ID"" type=""core:ContentTypeId""  "
                    + @"                                   use=""required"" />"
                    + @"                      <s:attribute name=""Description"" type=""s:string"" "
                    + @"                                   use=""required"" />"
                    + @"                      <s:attribute name=""Scope"" type=""s:string"" use=""required"" />"
                    + @"                      <s:attribute name=""Version"" type=""s:int"" use=""required"" />"
                    + @"                      <s:attribute name=""BestMatch"" type=""tns:TRUEONLY"" "
                    + @"                                   use=""optional"" />"
                    + @"                    </s:complexType>"
                    + @"                  </s:element>"
                    + @"                </s:sequence>"
                    + @"                <s:attribute name=""ContentTypeOrder"" type=""s:string"" use=""optional"" />"
                    + @"              </s:complexType>"
                    + @"            </s:element>"
                    + @"          </s:sequence>"
                    + @"        </s:complexType>"
                    + @"      </s:element>"
                    + @"    </s:sequence>"
                    + @"  </s:complexType>"
                    + @"</s:element>");

                // Verify R1835
                // Ensure the SOAP result is de-serialized successfully.
                Site.Assume.IsNotNull(getListContentTypesResult, "The result of GetListContentTypes operation must not be null.");

                // If the returned ContentType elements are not null, then the following requirement can be captured.
                bool isVerifyR1835 = true;
                foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in getListContentTypesResult.ContentTypes.ContentType)
                {
                    if (ct == null)
                    {
                        isVerifyR1835 = false;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1835,
                    1835,
                    @"[GetListContentTypesResponse]"
                    + "GetListContentTypesResult.ContentTypes.ContentType: The container element "
                    + "for a single content type, as specified in [MS-WSSCAML] section 2.4.");

                // Verify R1836
                // The response have been received successfully, then the following requirements can be captured.
                // If the response is not received and parsed successfully, the test case will fail before the requirements are captured 
                Site.CaptureRequirement(
                    1836,
                    @"[GetListContentTypesResponse]GetListContentTypesResult.ContentTypes.ContentType.XmlDocuments: 
                    A collection of XML documents associated with this content type, which can include forms and event receiver manifests.");

                // Verify R1837
                // The response have been received successfully, then the following requirements can be captured.
                // If the response is not received and parsed successfully, the test case will fail before the requirements are captured 
                Site.CaptureRequirement(
                    1837,
                    @"[GetListContentTypesResponse]The XmlDocumentDefinitionCollection type is " +
                    "specified in [MS-WSSCAML] section 2.4.12.");

                // Verify R1843
                // If the count of the returned BestMatch elements is not greater than 1, 
                // then the following requirement can be captured.
                int bestMatchCount = 0;
                foreach (GetListContentTypesResponseGetListContentTypesResultContentTypesContentType ct in getListContentTypesResult.ContentTypes.ContentType)
                {
                    if (ct.BestMatchSpecified == true)
                    {
                        bestMatchCount++;
                    }
                }

                bool isVerifyR1843 = bestMatchCount <= 1;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1843,
                    1843,
                    @"[GetListContentTypesResponse]"
                    + "GetListContentTypesResult.ContentTypes.ContentType.BestMatch: MUST be "
                    + "specified on at most one ContentType element.");

                // Verify the requirements of the TRUEONLY simple type.
                for (int i = 0; i < getListContentTypesResult.ContentTypes.ContentType.Length; i++)
                {
                    // Get the best match ContentTypeId, if BestMatchSpecified is true then we get it.
                    if (getListContentTypesResult.ContentTypes.ContentType[i].BestMatchSpecified)
                    {
                        this.VerifyTRUEONLY();
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListContentTypesAndProperties operation when the 
        /// response is received successfully.
        /// </summary>
        /// <param name="getListContentTypesAndPropertiesResult">The result of the operation.</param>
        private void VerifyGetListContentTypesAndPropertiesOperation(
            GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResult getListContentTypesAndPropertiesResult)
        {
            if (getListContentTypesAndPropertiesResult != null)
            {
                // Verify R1021
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1021,
                    @"[The schema of GetListContentTypesAndProperties is defined as: ]"
                    + @"<wsdl:operation name=""GetListContentTypesAndProperties"">"
                    + @"    <wsdl:input message=""GetListContentTypesAndPropertiesSoapIn"" />"
                    + @"    <wsdl:output message=""GetListContentTypesAndPropertiesSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R1023
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1023,
                    @"[In GetListContentTypesAndProperties operation]the protocol server responds "
                    + "with a GetListContentTypesAndPropertiesSoapOut message.");

                // Verify R1030
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1030,
                    @"[In GetListContentTypesAndPropertiesSoapOut]The SOAP Body contains a "
                    + "GetListContentTypesAndPropertiesResponse element.");

                // Verify R1056
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1056,
                    @"[The schema of GetListContentTypesAndPropertiesResponse is defined as:]"
                    + @"<s:element name=""GetListContentTypesAndPropertiesResponse"">"
                    + @"  <s:complexType>"
                    + @"    <s:sequence>"
                    + @"      <s:element name=""GetListContentTypesAndPropertiesResult"" minOccurs=""0"">"
                    + @"        <s:complexType mixed=""true"">"
                    + @"          <s:sequence>"
                    + @"            <s:element name=""ContentTypes"" >"
                    + @"              <s:complexType>"
                    + @"                <s:sequence>"
                    + @"                  <s:element name=""ContentType"" maxOccurs=""unbounded"">"
                    + @"                    <s:complexType>"
                    + @"                      <s:sequence>"
                    + @"                        <s:element name=""XmlDocuments"" "
                    + @"                                   type=""core:XmlDocumentDefinitionCollection"" "
                    + @"                                   minOccurs=""0"">"
                    + @"                        </s:element>"
                    + @"                      </s:sequence>"
                    + @"                      <s:attribute name=""Name"" type=""s:string"" use=""required"" />"
                    + @"                      <s:attribute name=""ID"" type=""core:ContentTypeId""  "
                    + @"                                   use=""required"" />"
                    + @"                      <s:attribute name=""Description"" type=""s:string"" "
                    + @"                                   use=""required"" />"
                    + @"                      <s:attribute name=""Scope"" type=""s:string"" use=""required"" />"
                    + @"                      <s:attribute name=""Version"" type=""s:int"" use=""required"" />"
                    + @"                      <s:attribute name=""BestMatch"" type=""tns:TRUEONLY"" "
                    + @"                                   use=""optional"" />"
                    + @"                    </s:complexType>"
                    + @"                  </s:element>"
                    + @"                </s:sequence>"
                    + @"                <s:attribute name=""ContentTypeOrder"" type=""s:string"" use=""optional"" />"
                    + @"              </s:complexType>"
                    + @"            </s:element>"
                    + @"            <s:element name=""Properties"">"
                    + @"              <s:complexType>"
                    + @"                <s:sequence>"
                    + @"                  <s:element name=""Property"" minOccurs=""0"" maxOccurs=""unbounded"">"
                    + @"                    <s:complexType>"
                    + @"                      <s:attribute name=""Key"" type=""s:string"" "
                    + @"                                   use=""required"" />"
                    + @"                      <s:attribute name=""Value"" type=""s:string"" "
                    + @"                                   use=""required"" />"
                    + @"                    </s:complexType>"
                    + @"                  </s:element>"
                    + @"                </s:sequence>"
                    + @"              </s:complexType>"
                    + @"            </s:element>"
                    + @"          </s:sequence>"
                    + @"        </s:complexType>"
                    + @"      </s:element>"
                    + @"    </s:sequence>"
                    + @"  </s:complexType>"
                    + @"</s:element>");

                // Verify R1059
                // Ensure the SOAP result is de-serialized successfully.
                Site.Assume.IsNotNull(getListContentTypesAndPropertiesResult, "The result of GetListContentTypesAndProperties operation must not be null.");

                bool isVerifyR1059 = true;
                foreach (
                    GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResultContentTypesContentType ct in getListContentTypesAndPropertiesResult.ContentTypes.ContentType)
                {
                    if (ct == null)
                    {
                        isVerifyR1059 = false;
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1059,
                    1059,
                    @"GetListContentTypesAndPropertiesResult.ContentTypes.ContentType: The "
                    + "container element for a single content type, as specified in [MS-WSSCAML] "
                    + "section 2.4.");

                // Verify R1060
                // The response have been received successfully, then the following requirements can be captured.
                // If the response is not received and parsed successfully, the test case will fail before the requirements are captured 
                Site.CaptureRequirement(
                    1060,
                    @"[In GetListContentTypesAndPropertiesResponse]GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.XmlDocuments: 
                    A collection of XML documents associated with this content type, which can include forms and event receiver manifests. ");

                // Verify R2235
                // The response have been received successfully, then the following requirements can be captured.
                // If the response is not received and parsed successfully, the test case will fail before the requirements are captured 
                Site.CaptureRequirement(
                    2235,
                    @"The XmlDocumentDefinitionCollection type is specified in [MS-WSSCAML] "
                    + "section 2.4.12");

                // Verify R1068
                // If the count of the returned BestMatch elements is not greater than 1, 
                // then the following requirement can be captured.
                int bestMatchCount = 0;
                foreach (GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResultContentTypesContentType
                    ct in getListContentTypesAndPropertiesResult.ContentTypes.ContentType)
                {
                    if (ct.BestMatchSpecified == true)
                    {
                        bestMatchCount++;
                    }
                }

                bool isVerifyR1068 = bestMatchCount <= 1;
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1068,
                    1068,
                    @"[In GetListContentTypesAndPropertiesResponse]"
                    + "GetListContentTypesAndPropertiesResult.ContentTypes.ContentType.BestMatch: "
                    + "MUST be specified on at most one ContentType element.");
            }

            // Verify the requirements of the TRUEONLY simple type.
            foreach (GetListContentTypesAndPropertiesResponseGetListContentTypesAndPropertiesResultContentTypesContentType ct in getListContentTypesAndPropertiesResult.ContentTypes.ContentType)
            {
                // Get the best match ContentTypeId, if BestMatchSpecified is true then we get it.
                if (ct.BestMatchSpecified)
                {
                    this.VerifyTRUEONLY();
                    break;
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListItemChanges operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="getListItemChangesResult">The result of the operation</param>
        private void VerifyGetListItemChangesOperation(
            GetListItemChangesResponseGetListItemChangesResult getListItemChangesResult)
        {
            Site.Assert.IsNotNull(getListItemChangesResult, "The getListItemChangesResult must be not null.");

            // Verify R1845
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1845,
                @"[The schema of GetListItemChanges is defined as:]"
                + @"<wsdl:operation name=""GetListItemChanges"">"
                + @"    <wsdl:input message=""GetListItemChangesSoapIn"" />"
                + @"    <wsdl:output message=""GetListItemChangesSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R628
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                628,
                @"[In GetListItemChanges operation] [If the protocol client sends a "
                + "GetListItemChangesSoapIn request message] the protocol server responds "
                + "with a GetListItemChangesSoapOut response message.");

            // Verify R1850
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured  
            Site.CaptureRequirement(
                1850,
                @"[GetListItemChangesSoapOut]The SOAP Body contains a "
                + "GetListItemChangesResponse element.");

            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(getListItemChangesResult, "The result of GetListItemChanges operation must not be null.");

            // Verify R1856
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1856,
                @"[The schema of GetListItemChangesResponse is defined as:] "
                + @"<s:element minOccurs=""0"" maxOccurs=""1"" name=""GetListItemChangesResult"">"
                + @"   <s:complexType mixed=""true"">"
                + @"    <s:sequence>"
                + @"      <s:element name=""GetListItemChangesResult"">"
                + @"        <s:complexType>"
                + @"          <s:sequence>"
                + @"            <s:element name=""listitems"" >"
                + @"              <s:complexType mixed=""true"" >"
                + @"                <s:sequence>"
                + @"                  <s:any />                "
                + @"                </s:sequence>"
                + @"              <s:Attribute name=""TimeStamp"" type=""s:string""/>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1857
            // When the index of the <s:schema> element is equal to -1, which 
            // means the element is excluded, then the following requirement can be captured
            this.Site.Assert.IsNotNull(getListItemChangesResult.listitems, "listitems element should not be null.");

            // There can be a maximum of two rs:data elements
            if (null != getListItemChangesResult.listitems.data && getListItemChangesResult.listitems.data.Length <= 2)
            {
                bool isunExpectedelementExisting = false;
                foreach (DataDefinition dataElemt in getListItemChangesResult.listitems.data)
                {
                    if (null != dataElemt.Any)
                    {
                        XmlNode[] rowdatas = dataElemt.Any;

                        // verify whether exiting a "<s:schema>" under listitems element
                        isunExpectedelementExisting = rowdatas.Any(founder => (founder.OuterXml.IndexOf(
                                                            "<s:schema>",
                                                             StringComparison.OrdinalIgnoreCase) >= 0));
                    }

                    if (isunExpectedelementExisting)
                    {
                        break;
                    }
                }

                // if there is no "<s:schema>" element existing, capture R1587
                Site.CaptureRequirementIfIsFalse(
                    isunExpectedelementExisting,
                    1857,
                    @"[GetListItemChangesResponse]GetListItemChangesResult: This protocol server "
                    + "response included in the listitems element is modeled on the Microsoft ADO "
                    + "2.6 Persistence format [MS-PRSTFR], excluding the <s:schema> element.");
            }

            // Verify R2163
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2163,
                @"[In GetListItemChanges operation] [In GetListItemChangesResponse element]"
                + "The listitems element includes attributes describing the namespaces for the "
                + "ADO 2.6 Persistence format.");

            // Verify R1858
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1858,
                @"[GetListItemChangesResponse] listitems contains an inner element named "
                + "rs:data, which is of type DataDefinition.");

            // Verify R1859
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            DateTime dt = new DateTime();
            bool isVerifyR1859 = DateTime.TryParse(getListItemChangesResult.listitems.TimeStamp, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dt);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1859,
                1859,
                @"[GetListItemChangesResponse]The TimeStamp attribute is a string that contains the "
                + "date in Coordinated Universal Time (UTC) of the request to the protocol server.");

            // Verify R1860
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured
            Site.CaptureRequirement(
                1860,
                @"[GetListItemChangesResponse]There can be a maximum of two rs:data elements.");

            // Verify R1864
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1864,
                @"[GetListItemChangesResponse]Note that set of fields returned by the method is "
                + "restricted by the viewField parameter.");
        }

        /// <summary>
        /// Verify the message syntax of GetListItemChangesSinceToken operation when the response 
        /// is received successfully.
        /// </summary>
        /// <param name="getListItemChangesSinceTokenResult">The result of the operation</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="viewFields">Specifies which fields of the list item will be returned</param>
        private void VerifyGetListItemChangesSinceTokenOperation(
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesSinceTokenResult,
            CamlQueryOptions queryOptions,
            CamlViewFields viewFields)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(getListItemChangesSinceTokenResult, "The result of GetListItemChangesSinceToken operation must not be null.");

            // Verify R1876
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1876,
                @"[The schema of GetListItemChangesSinceToken is defined as:]"
                + @"<wsdl:operation name=""GetListItemChangesSinceToken"">"
                + @"    <wsdl:input message=""GetListItemChangesSinceTokenSoapIn"" />"
                + @"    <wsdl:output message=""GetListItemChangesSinceTokenSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R654
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                654,
                @"[In GetListItemChangesSinceToken operation] [If the protocol client sends a "
                + "GetListItemChangesSinceTokenSoapIn request message] the protocol server "
                + "responds with a GetListItemChangesSinceTokenSoapOut response message.");

            // Verify R1882
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1882,
                @"[GetListItemChangesSinceTokenSoapOut]The SOAP Body contains a "
                + "GetListItemChangesSinceTokenResponse element.");

            // Verify R1897
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1897,
                @"[The schema of GetListItemChangesSinceTokenResponse is defined as:] <s:element name=""GetListItemChangesSinceTokenResponse"">
              <s:complexType>
                <s:sequence>
                    <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetListItemChangesSinceTokenResult"">
                     <s:complexType mixed=""true"">
                      <s:sequence>
                        <s:element name=""listitems"" >
                          <s:complexType>
                            <s:sequence>
                              <s:element name=""Changes"" >
                                <s:complexType>
                                  <s:sequence>
                                    <s:element name=""List"" type=""tns:ListDefinitionSchema""  
                                               minOccurs=""0"" />
                                    <s:element name=""Id"" type=""tns:ListItemChangeDefinition"" minOccurs=""0""/>                      
                                  </s:sequence>
                                  <s:attribute name=""LastChangeToken"" type=""s:string"" />
                                  <s:attribute name=""MoreChanges"" type=""core:TRUEFALSE"" />
                                  <s:attribute name=""MinTimeBetweenSyncs"" type=""s:unsignedInt"" />
                                  <s:attribute name=""RecommendedTimeBetweenSyncs"" type=""s:unsignedInt"" />
                                  <s:attribute name=""MaxBulkDocumentSyncSize"" type=""s:unsignedInt"" />
                                  <s:attribute name=""MaxRecommendedEmbeddedFileSize"" type=""s:unsignedInt"" />
                                  <s:attribute name=""AlternateUrls"" type=""s:string"" />
                                  <s:attribute name=""EffectivePermMask"" type=""s:string"" />
                                </s:complexType>
                              </s:element>
                              <s:any />
                            </s:sequence> 
                          </s:complexType>
                        </s:element>
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                </s:sequence>     
              </s:complexType>
            </s:element>");

            // Verify R1898
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1898,
                @"[GetListItemChangesSinceTokenResponse]GetListItemChangesSinceTokenResult:  The top-level element, which contains a listitems element.");

            // Verify R1908
            // In MS-LISTSWS.wsdl. The 'any' element has been replaced by rs:data.
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1908,
                @"[GetListItemChangesSinceTokenResponse]The listitems element also contains an inner element named rs:data, which is of type DataDefinition ListItemCollectionPositionNext (section 2.2.4.7).");

            // Verify R2486
            if (Common.IsRequirementEnabled(2486, this.Site))
            {
                if (getListItemChangesSinceTokenResult.listitems.Changes.List != null)
                {
                    bool isFileFragmentExist = getListItemChangesSinceTokenResult.listitems.Changes.List.Fields.Field.Any(field => field.Name == "FileFragment");
                    Site.CaptureRequirementIfIsFalse(
                        isFileFragmentExist,
                        2486,
                        @"[In GetListItemChangesSinceToken operation]Implementation does not return the FileFragment element.[In Appendix B: Product Behavior] <70> Section 3.1.4.22.2.2: In Windows SharePoint Services 3.0, the FileFragment element is not returned.");
                }
            }

            // Verify the requirements of the DataDefinition complex type.
            if (getListItemChangesSinceTokenResult.listitems.data != null)
            {
                this.VerifyDataDefinition(getListItemChangesSinceTokenResult.listitems.data);
            }

            // Verify the requirements of the ListDefinitionSchema complex type.
            if (getListItemChangesSinceTokenResult.listitems.Changes.List != null)
            {
                this.VerifyListDefinitionSchema(getListItemChangesSinceTokenResult.listitems.Changes.List);
            }

            // Verify the requirements of the ListItemChangeDefinition complex type.
            if (null != getListItemChangesSinceTokenResult.listitems && null != getListItemChangesSinceTokenResult.listitems.Changes
                && null != getListItemChangesSinceTokenResult.listitems.Changes && null != getListItemChangesSinceTokenResult.listitems.Changes.Id)
            {
                this.VerifyListItemChangeDefinition(getListItemChangesSinceTokenResult.listitems.Changes.Id);
                this.VerifyServerChangeUnitAttributeNotReturn(getListItemChangesSinceTokenResult.listitems.Changes.Id);
            }

            // Verify the requirements of the EnumViewAttributes simple type.
            if (queryOptions != null)
            {
                if (getListItemChangesSinceTokenResult.listitems.data.Any != null)
                {
                    DataTable data = AdapterHelper.ExtractData(getListItemChangesSinceTokenResult.listitems.data.Any);
                    string author = data.Columns.Contains("ows_Author") ? Convert.ToString(data.Rows[0]["ows_Author"]) : null;
                    this.VerifyCamlQueryOptions(queryOptions, viewFields, author);
                }

                if (queryOptions.QueryOptions != null)
                {
                    if (queryOptions.QueryOptions.ViewAttributes != null)
                    {
                        if (queryOptions.QueryOptions.ViewAttributes.ScopeSpecified)
                        {
                            this.VerifyEnumViewAttributes();
                        }
                    }
                }
            }
            // Verify R1907
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                   1907,
                   @"[GetListItemChangesSinceTokenResponse]Note that set of fields returned by the method is restricted by the viewField or viewName parameter.");
        }

        /// <summary>
        /// Verify the message syntax of GetListItemChangesWithKnowledge operation when the 
        /// response is received successfully.
        /// </summary>
        /// <param name="getListItemChangesWithKnowledgeResult">The result of the operation</param>
        /// <param name="queryOptions">Specifies various options for modifying the query</param>
        /// <param name="viewFields">Specifies which fields of the list item will be returned</param>
        private void VerifyGetListItemChangesWithKnowledgeOperation(
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListItemChangesWithKnowledgeResult,
            CamlQueryOptions queryOptions,
            CamlViewFields viewFields)
        {
            if (getListItemChangesWithKnowledgeResult != null)
            {
                // Verify R1079
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1079,
                    @"[The schema of GetListItemChangesWithKnowledge is defined as:]"
                    + @"<wsdl:operation name=""GetListItemChangesWithKnowledge"">"
                    + @"    <wsdl:input message=""GetListItemChangesWithKnowledgeSoapIn"" />"
                    + @"    <wsdl:output message=""GetListItemChangesWithKnowledgeSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R1081
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1081,
                    @"[In GetListItemChangesWithKnowledge operation]The protocol client sends a "
                    + "GetListItemChangesWithKnowledgeSoapIn request message ,the protocol "
                    + "server responds with a GetListItemChangesWithKnowledgeSoapOut response "
                    + "message,");

                // Verify R1096
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1096,
                    @"[In GetListItemChangeswithKnowledgeSoapOut]The SOAP Body contains a "
                    + "GetListItemChangesWithKnowledgeResponse element.");

                // Verify R1124
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1124,
                    @"[The schema of GetListItemChangesWithKnowledgeResponse is defined as: ]
                <s:element name=""GetListItemChangesWithKnowledgeResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetListItemChangesWithKnowledgeResult"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""listitems"" >
                              <s:complexType>
                                <s:sequence>
                                  <s:element name=""Changes"" >
                                    <s:complexType>
                                      <s:sequence>
                                        <s:element name=""MadeWithKnowledge"" minOccurs=""0"" maxOccurs=""1"">
                                          <s:complexType>
                                            <s:sequence>
                                              <s:element ref=""sync:syncKnowledge"" />
                                            </s:sequence>
                                          </s:complexType>
                                        </s:element>
                                        <s:element name=""Id"" type=""tns:ListItemChangeDefinition"" minOccurs=""0""/>
                                        <s:element name=""View"" type=""tns:ViewChangeDefinition"" minOccurs=""0""/>
                                      </s:sequence>
                                      <s:attribute name=""SchemaChanged"" type=""core:TRUEFALSE"" />
                                      <s:attribute name=""ServerTime"" type=""s:string"" />
                                      <s:attribute name=""MoreChanges"" type=""core:TRUEFALSE"" />
                                      <s:attribute name=""MinTimeBetweenSyncs"" type=""s:unsignedInt"" />
                                      <s:attribute name=""RecommendedTimeBetweenSyncs"" type=""s:unsignedInt"" />
                                      <s:attribute name=""MaxBulkDocumentSyncSize"" type=""s:unsignedInt"" />
                                      <s:attribute name=""MaxRecommendedEmbeddedFileSize"" type=""s:unsignedInt"" />
                                      <s:attribute name=""AlternateUrls"" type=""s:string"" />
                                      <s:attribute name=""EffectivePermMask"" type=""s:string"" />
                                    </s:complexType>
                                  </s:element>
                                  <s:any />
                                </s:sequence>
                              </s:complexType>
                            </s:element>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>     
                  </s:complexType>
                </s:element>");

                // Verify R1125
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
                Site.CaptureRequirement(
                    1125,
                    @"[In GetListItemChangesWithKnowledgeResponse]GetListItemChangesWithKnowledgeResult: The top-level element, which contains a listitems element.");

                // Verify R1137
                // The schema from MS-XSSK has been added to MS-LISTSWS.wsdl.
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
                Site.CaptureRequirement(
                    1137,
                    @"[In GetListItemChangesWithKnowledgeResponse]The inner XML of the MadeWithKnowledge element in the Changes element is the knowledge in XML format, as specified in [MS-XSSK] section 3, that represents the last change in the list that is returned to the client.");

                // Verify R1145
                // The response have been received successfully, which means the schema
                // of the listitems element past the validation, then the following requirement can be 
                // captured.
                Site.CaptureRequirement(
                    1145,
                    @"[In GetListItemChangesWithKnowledgeResponse]The listitems element also "
                    + "contains an inner element named rs:data, which is of type DataDefinition. (section 2.2.4.7)");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R1361
                // If the Created can be parsed to a DateTime, capture R1135
                if (null != getListItemChangesWithKnowledgeResult.listitems && null != getListItemChangesWithKnowledgeResult.listitems.Changes
                    && !string.IsNullOrEmpty(getListItemChangesWithKnowledgeResult.listitems.Changes.ServerTime))
                {
                    DateTime created;
                    string paserFormat = @"yyyyMMdd HH:mm:ss";
                    string returnDataTimeValue = getListItemChangesWithKnowledgeResult.listitems.Changes.ServerTime;
                    bool isVerifyR1135 = DateTime.TryParseExact(returnDataTimeValue, paserFormat, CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out created);

                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual values: returnDataTimeValue[{0}] for requirement #R1361",
                        string.IsNullOrEmpty(returnDataTimeValue) ? "NullOrEmpty" : returnDataTimeValue);

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR1135,
                        1135,
                        @"[In GetListItemChangesWithKnowledgeResponse][Attribute of ServerTime]The ServerTime attribute specifies the UTC date and time in the Gregorian calendar "
                        + @"when the changes were returned by the server in the format ""yyyyMMdd hh:mm:ss"" where ""yyyy"" represents the year, ""MM"" represents the month, ""dd"" represents the day of the month, "
                        + @"""hh"" represents the hour, ""mm"" represents the minute, and ""ss"" represents the second.");
                }
            }

            if ((queryOptions != null)
                && (getListItemChangesWithKnowledgeResult.listitems.data.Any != null))
            {
                DataTable data = AdapterHelper.ExtractData(getListItemChangesWithKnowledgeResult.listitems.data.Any);
                string author = data.Columns.Contains("ows_Author") ? Convert.ToString(data.Rows[0]["ows_Author"]) : null;
                this.VerifyCamlQueryOptions(queryOptions, viewFields, author);
            }

            // Verify the requirements of the DataDefinition complex type.
            if (getListItemChangesWithKnowledgeResult.listitems.data != null)
            {
                this.VerifyDataDefinition(getListItemChangesWithKnowledgeResult.listitems.data);
            }

            // Verify the requirements of FileFolderChangeDefinition complex type.
            if (getListItemChangesWithKnowledgeResult.listitems.Changes.File != null)
            {
                this.VerifyFileFolderChangeDefinition(getListItemChangesWithKnowledgeResult.listitems.Changes.File);
            }

            if (getListItemChangesWithKnowledgeResult.listitems.Changes.Folder != null)
            {
                this.VerifyFileFolderChangeDefinition(getListItemChangesWithKnowledgeResult.listitems.Changes.Folder);
            }

            // Verify the requirements of FileFragmentChangeDefinition complex type.
            if (getListItemChangesWithKnowledgeResult.listitems.Changes.FileFragment != null)
            {
                this.VerifyFileFragmentChangeDefinition(getListItemChangesWithKnowledgeResult.listitems.Changes.FileFragment);
            }

            // Verify the requirements of the ListItemChangeDefinition complex type.
            if (getListItemChangesWithKnowledgeResult.listitems.Changes.Id != null)
            {
                this.VerifyListItemChangeDefinition(getListItemChangesWithKnowledgeResult.listitems.Changes.Id[0]);
            }

            // Verify the requirements of the EnumViewAttributes simple type.
            if (queryOptions != null)
            {
                if (queryOptions.QueryOptions != null)
                {
                    if (queryOptions.QueryOptions.ViewAttributes != null)
                    {
                        if (queryOptions.QueryOptions.ViewAttributes.ScopeSpecified)
                        {
                            this.VerifyEnumViewAttributes();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of GetListItems operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="getListItemResult">The result of the operation.</param>
        /// <param name="queryOptions">Specifies various options for modifying the query.</param>
        /// <param name="viewFields">Specifies which fields of the list item will be returned.</param>
        private void VerifyGetListItemsOperation(GetListItemsResponseGetListItemsResult getListItemResult, CamlQueryOptions queryOptions, CamlViewFields viewFields)
        {
            // Verify R1910
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1910,
                @"[The schema of GetListItems is defined as:]"
                + @"<wsdl:operation name=""GetListItems"">"
                + @"    <wsdl:input message=""GetListItemsSoapIn"" />"
                + @"    <wsdl:output message=""GetListItemsSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R723
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                723,
                @"[In GetListItems operation] [If the protocol client sends a GetListItemsSoapIn "
                + "request message] the protocol server responds with a GetListItemsSoapOut "
                + "response message.");

            // Verify R1918
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1918,
                @"[GetListItemsSoapOut]The SOAP Body contains a GetListItemsResponse "
                + "element.");

            // Verify R1935
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1935,
                @"[The schema of GetListItemsResponse is defined as:] "
                + @"<s:element name=""GetListItemsResponse"">"
                + @"  <s:complexType mixed=""true"">"
                + @"    <s:sequence>"
                + @"      <s:element minOccurs=""0"" maxOccurs=""1"" name=""GetListItemsResult"">"
                + @"        <s:complexType>"
                + @"          <s:sequence>"
                + @"            <s:element name=""listitems"" >"
                + @"              <s:complexType mixed=""true"" >"
                + @"                <s:sequence>"
                + @"                  <s:any />"
                + @"                </s:sequence>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1936
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1936,
                @"[GetListItemsResponse]GetListItemsResult: This protocol server response included in "
                + "the listitems element is modeled on the Microsoft ADO 2.6 Persistence format "
                + "[MS-PRSTFR], excluding the <s:schema> element.");

            // Verify R2330
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2330,
                @"[In GetListitems operation] [In GetListitemsResponse element] [In GetListItemsResult "
                + "element]The listitems element includes attributes describing the namespaces "
                + "for the ADO 2.6 Persistence format.");

            // Verify R2331
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2331,
                @"[In GetListitems operation] [In GetListitemsResponse element] [In GetListItemsResult "
                + "element] [The listitems element] contains an inner element named rs:data, which is "
                + "of type DataDefinition.");

            // Verify R1938
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1938,
                @"[GetListItemsResponse]Note that set of fields returned by the method is restricted "
                + "by the viewField or viewName parameter.");

            // Verify R1940
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1940,
                @"[GetListItemsResponse]The listitems element contains attributes that define the "
                + "namespaces.");

            // Verify R1941
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1941,
                @"[GetListItemsResponse]Inside of this element[listitems] is the <rs:data> element, which "
                + "specifies how many rows of data are being returned, where a row of data "
                + "corresponds to a list item, and the paging token (if there are more rows in "
                + "the view than were returned).");

            // Verify the requirements of the EnumViewAttributes simple type.
            if (queryOptions != null)
            {
                if (getListItemResult.listitems.data.Any != null)
                {
                    DataTable data = AdapterHelper.ExtractData(getListItemResult.listitems.data.Any);
                    string author = data.Columns.Contains("ows_Author") ? Convert.ToString(data.Rows[0]["ows_Author"]) : null;
                    this.VerifyCamlQueryOptions(queryOptions, viewFields, author);
                }

                if (queryOptions.QueryOptions != null)
                {
                    if (queryOptions.QueryOptions.ViewAttributes != null)
                    {
                        if (queryOptions.QueryOptions.ViewAttributes.ScopeSpecified)
                        {
                            this.VerifyEnumViewAttributes();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of GetVersionCollection operation when the response is 
        /// received successfully.
        /// </summary>
        /// <param name="getVersionCollectionResult">The result of the operation</param>
        private void VerifyGetVersionCollectionOperation(GetVersionCollectionResponseGetVersionCollectionResult getVersionCollectionResult)
        {
            Site.Assert.IsNotNull(
                getVersionCollectionResult,
                "The return value of this operation cannot be null.");

            // Verify R1942
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1942,
                @"[The schema of GetVersionCollection is defined as:]"
                + @"<wsdl:operation name=""GetVersionCollection"">"
                + @"    <wsdl:input message=""GetVersionCollectionSoapIn"" />"
                + @"    <wsdl:output message=""GetVersionCollectionSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R761
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                761,
                @"[In GetVersionCollection operation] [If the protocol client sends a "
                + "GetVersionCollectionSoapIn request message] the protocol server responds with "
                + "a GetVersionCollectionSoapOut response message.");

            // Verify R1950
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1950,
                @"[GetVersionCollectionSoapOut]The SOAP Body contains a "
                + "GetVersionCollectionResponse element.");

            // Verify R776
            // If the returned version count is greater than 0, then the following
            // requirements can be captured.
            bool isVerifyR776 = getVersionCollectionResult.Versions.Length > 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR776,
                776,
                @"[In GetVersionCollection operation] If the operation succeeds, the collection "
                + "of versions MUST be returned for the specified field.");

            // Verify R1955
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1955,
                @"[GetVersionCollectionResponse]"
                + @"<s:element name=""GetVersionCollectionResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""GetVersionCollectionResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Versions"">"
                + @"              <s:complexType>"
                + @"                <s:sequence>"
                + @"                  <s:element name=""Version"" minOccurs=""0"" "
                + @"maxOccurs=""unbounded"">"
                + @"                    <s:complexType>"
                + @"                      <s:attribute name=""FieldName"" type=""s:string""/>"
                + @"                      <s:attribute name=""Modified"" type=""s:string""/>"
                + @"                      <s:attribute name=""Editor"" type=""s:string""/>"
                + @"                    </s:complexType>"
                + @"                  </s:element>"
                + @"                </s:sequence>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1957
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1957,
                @"[GetVersionCollectionResponse]Versions: The collection of versions for the specified list item.");
        }

        /// <summary>
        /// Verify the message syntax of UndoCheckOut operation when the response is received 
        /// successfully.
        /// </summary>
        private void VerifyUndoCheckOutOperation()
        {
            // Verify R1961
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1961,
                @"[The schema of UndoCheckOut is defined as:]"
                + @"<wsdl:operation name=""UndoCheckOut"">"
                + @"    <wsdl:input message=""UndoCheckOutSoapIn"" />"
                + @"    <wsdl:output message=""UndoCheckOutSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R782
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                782,
                @"[In UndoCheckOut operation] [If the protocol client sends an UndoCheckOutSoapIn "
                + "request message] the protocol server responds with an UndoCheckOutSoapOut "
                + "response message.");

            // Verify R1968
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1968,
                @"[UndoCheckOutSoapOut]The SOAP Body contains an UndoCheckOutResponse "
                + "element.");

            // Verify R1971
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1971,
                @"[The schema of UndoCheckOutResponse is defined as:]"
                + @"<s:element name=""UndoCheckOutResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""UndoCheckOutResult"" type=""s:boolean""/>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1972
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1972,
                @"[UndoCheckOutResponse]UndoCheckOutResult: Specifies whether the call is "
                + "successful or not.");
        }

        /// <summary>
        /// Verify the message syntax of UpdateContentType operation when the response is received successfully.
        /// </summary>
        /// <param name="updateContentTypeResult">The result of the UpdateContentType operation</param>
        private void VerifyUpdateContentTypeOperation(UpdateContentTypeResponseUpdateContentTypeResult updateContentTypeResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateContentTypeResult,
                "The result of UpdateContentType operation must not be null.");

            // Verify R1974
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1974,
                @"[The schema of UpdateContentType is defined as:]"
                + @"<wsdl:operation name=""UpdateContentType"">"
                + @"    <wsdl:input message=""UpdateContentTypeSoapIn"" />"
                + @"    <wsdl:output message=""UpdateContentTypeSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R796
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured
            Site.CaptureRequirement(
                796,
                @"[In UpdateContentType operation] [If the protocol client sends an "
                + "UpdateContentTypeSoapIn request message] the protocol server responds with "
                + "an UpdateContentTypeSoapOut response message.");

            // Verify R1983
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1982,
                @"[UpdateContentTypeSoapOut]The SOAP action value of the message is defined as follows:
                http://schemas.microsoft.com/sharepoint/soap/UpdateContentType");

            // Verify R1983
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1983,
                @"[UpdateContentTypeSoapOut]The SOAP body contains an "
                + @"UpdateContentTypeResponse element.");

            // Verify R1993
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1993,
                @"[UpdateContentTypeResponse] "
                + @"<s:element name=""UpdateContentTypeResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element minOccurs=""0"" maxOccurs=""1"" name=""UpdateContentTypeResult"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Results"">"
                + @"              <s:complexType>"
                + @"                <s:sequence>"
                + @"                  <s:element name=""Method"" minOccurs=""0"" "
                + @"maxOccurs=""unbounded"">"
                + @"                    <s:complexType>"
                + @"                      <s:sequence>"
                + @"                        <s:element name=""ErrorCode"" type=""s:string"" />"
                + @"                        <s:element name=""FieldRef"" "
                + @"                                   type=""tns:FieldReferenceDefinitionCT"" "
                + @"                                   minOccurs=""0"" />"
                + @"                        <s:element name=""Field"" type=""core:FieldDefinition"" "
                + @"                                   minOccurs=""0"" />"
                + @"                        <s:element name=""ErrorText"" "
                + @"                                   type=""s:string"" minOccurs=""0"" />"
                + @"                      </s:sequence>"
                + @"                      <s:attribute name=""ID"" type=""s:string"" use=""required""/>"
                + @"                    </s:complexType>"
                + @"                  </s:element>"
                + @"                  <s:element name=""ListProperties"">"
                + @"                    <s:complexType>"
                + @"                      <s:attribute name=""Description"" type=""s:string"" "
                + @"                                   use=""optional"" />"
                + @"                      <s:attribute name=""FeatureId"" "
                + @"                                   type=""core:UniqueIdentifierWithOrWithoutBraces"" "
                + @"                                   use=""optional""/>"
                + @"                      <s:attribute name=""Group"" type=""s:string"" use=""optional"" />"
                + @"                      <s:attribute name=""Hidden"" "
                + @"                                   type=""core:TRUE_NegOne_Else_Anything"" "
                + @"                                   use=""optional"" />"
                + @"                      <s:attribute name=""ID"" type=""core:ContentTypeId"" "
                + @"                                   use=""required"" />"
                + @"                      <s:attribute name=""Name"" type=""s:string"" use=""required"" />"
                + @"                      <s:attribute name=""ReadOnly"" "
                + @"                                   type=""core:TRUE_NegOne_Else_Anything"" "
                + @"                                   use=""optional"" />"
                + @"                      <s:attribute name=""Sealed"" "
                + @"                                   type=""core:TRUE_Case_Sensitive_Else_Anything"" "
                + @"                                   use=""optional"" />"
                + @"                      <s:attribute name=""V2ListTemplateName"" type=""s:string"" "
                + @"                                   use=""optional""/>"
                + @"                      <s:attribute name=""Version"" type=""s:long"" use=""optional"" />"
                + @"                      <s:anyAttribute namespace=""##other"" processContents=""lax"" />"
                + @"                    </s:complexType>"
                + @"                  </s:element>"
                + @"                </s:sequence>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify R1994
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                1994,
                @"[UpdateContentTypeResponse]Results: The container for data on the update of a content type.");

            // Verify R1995
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                1995,
                @"[UpdateContentTypeResponse]Method: The container for data on a field add, update, or remove operation.");

            // Verify R1997
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                1997,
                @"[UpdateContentTypeResponse]Field: A FieldDefinition, as specified by [MS-WSSFO2] "
                + "section 2.2.8.3.3.[A field definition describes the structure and format of a field that "
                + "is used within a list or content type.]");

            // Verify R1998
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                1998,
                @"[UpdateContentTypeResponse]ListProperties: Contains updated data for the content type.");

            // Verify the requirements of FieldReferenceDefinitionCT complex type.
            foreach (UpdateContentTypeResponseUpdateContentTypeResultResultsMethod method in updateContentTypeResult.Results.Method)
            {
                if (method.FieldRef != null)
                {
                    this.VerifyFieldReferenceDefinitionCT(method.FieldRef);
                }
            }
        }

        /// <summary>
        /// Verify the message syntax of UpdateContentTypesXmlDocument operation when the response 
        /// is received successfully.
        /// </summary>
        /// <param name="updateContentTypesXmlDocumentResult">The result of the operation</param>
        private void VerifyUpdateContentTypesXmlDocumentOperation(UpdateContentTypesXmlDocumentResponseUpdateContentTypesXmlDocumentResult updateContentTypesXmlDocumentResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateContentTypesXmlDocumentResult,
                "The result of UpdateContentTypesXmlDocument operation must not be null.");

            // Verify R2008
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2008,
                @"[The schema of UpdateContentTypesXmlDocument is defined as:]"
                + @"<wsdl:operation name=""UpdateContentTypesXmlDocument"">"
                + @"    <wsdl:input message=""UpdateContentTypesXmlDocumentSoapIn"" />"
                + @"    <wsdl:output message=""UpdateContentTypesXmlDocumentSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R839
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                839,
                @"[In UpdateContentTypesXmlDocument operation] [If the protocol client sends an "
                + "UpdateContentTypesXmlDocumentSoapIn request message] the protocol server "
                + "responds with an UpdateContentTypesXmlDocumentSoapOut response message.");

            // Verify R2015
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2015,
                @"[UpdateContentTypesXmlDocumentSoapOut]The SOAP Body contains an "
                + @"UpdateContentTypesXmlDocumentResponse element.");

            // Verify R2023
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2023,
                @"[UpdateContentTypesXmlDocumentResponse]"
                + @"<s:element name=""UpdateContentTypesXmlDocumentResponse"">"
                + @"  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""UpdateContentTypesXmlDocumentResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Success"" minOccurs=""0"">"
                + @"              <s:complexType />"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");
        }

        /// <summary>
        /// Verify the message syntax of UpdateContentTypeXmlDocument operation when the response 
        /// is received successfully.
        /// </summary>
        /// <param name="updateContentTypeXmlDocumentResult">The result of the operation.</param>
        private void VerifyUpdateContentTypeXmlDocumentOperation(System.Xml.XmlNode updateContentTypeXmlDocumentResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateContentTypeXmlDocumentResult,
                "The result of UpdateContentTypeXmlDocument operation must not be null.");

            // Verify R2025
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2025,
                @"[The schema of UpdateContentTypeXmlDocument is defined as:]"
                + @"<wsdl:operation name=""UpdateContentTypeXmlDocument"">"
                + @"    <wsdl:input message=""UpdateContentTypeXmlDocumentSoapIn"" />"
                + @"    <wsdl:output message=""UpdateContentTypeXmlDocumentSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R862
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                862,
                @"[In UpdateContentTypeXmlDocument operation] [If the protocol client sends "
                + "an UpdateContentTypeXmlDocumentSoapIn request message] the protocol server "
                + "responds with an UpdateContentTypeXmlDocumentSoapOut response message.");

            // Verify R870
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateContentTypeXmlDocumentResult,
                "The result of UpdateContentTypeXmlDocument operation must not be null.");

            // If no exception thrown and the result is not null, then the following 
            // requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                updateContentTypeXmlDocumentResult,
                870,
                @"[In UpdateContentTypeXmlDocument operation] If no SOAP fault is thrown, the "
                + "protocol server MUST return a success UpdateContentTypeXmlDocumentResult.");

            // Verify R2031
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2031,
                @"[UpdateContentTypeXmlDocumentSoapOut]The SOAP Body contains an "
                + "UpdateContentTypeXmlDocumentResponse element.");

            // Verify R2036
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2036,
                @"[UpdateContentTypeXmlDocumentResponse]: <s:element name=""UpdateContentTypeXmlDocumentResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element name=""UpdateContentTypeXmlDocumentResult"" minOccurs=""0"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                             <s:element name=""Success"" minOccurs=""0"">
                            <s:complexType />
                            </s:element>
                          </s:sequence>
                        </s:complexType>
                      </s:element>
                    </s:sequence>
                  </s:complexType>
                </s:element>");
        }

        /// <summary>
        /// Verify the message syntax of UpdateList operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="updateListResult">The result of the operation.</param>
        private void VerifyUpdateListOperation(UpdateListResponseUpdateListResult updateListResult)
        {
            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateListResult,
                "The result of UpdateList operation must not be null.");

            // Verify R2038
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2038,
                @"[The schema of UpdateList is defined as:]<wsdl:operation name=""UpdateList"">"
                + @"    <wsdl:input message=""UpdateListSoapIn"" />"
                + @"    <wsdl:output message=""UpdateListSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R883
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                883,
                @"[In UpdateList operation] [If the protocol client sends an UpdateListSoapIn "
                + "request message] The server MUST respond with an UpdateListSoapOut response "
                + "message.");

            // Verify R2046
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2046,
                @"[UpdateListSoapOut]The SOAP Body contains an UpdateListResponse element.");

            // Verify R2061
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2061,
                @"[The schema of UpdateListResponse is defined as:]"
                + @"<s:element name=""UpdateListResponse"">  <s:complexType>"
                + @"    <s:sequence>"
                + @"      <s:element name=""UpdateListResult"" minOccurs=""0"">"
                + @"        <s:complexType mixed=""true"">"
                + @"          <s:sequence>"
                + @"            <s:element name=""Results"">"
                + @"              <s:complexType mixed=""true"">"
                + @"                <s:sequence>"
                + @"                  <s:element name=""NewFields"" type=""tns:UpdateListFieldResults"" />"
                + @"                  <s:element name=""UpdateFields"" "
                + @"                             type=""tns:UpdateListFieldResults"" />"
                + @"                  <s:element name=""DeleteFields"" "
                + @"                             type=""tns:UpdateListFieldResults"" />"
                + @"                  <s:element name=""ListProperties"" type=""tns:ListDefinitionCT"" />"
                + @"                </s:sequence>"
                + @"              </s:complexType>"
                + @"            </s:element>"
                + @"          </s:sequence>"
                + @"        </s:complexType>"
                + @"      </s:element>"
                + @"    </s:sequence>"
                + @"  </s:complexType>"
                + @"</s:element>");

            // Verify the requirements of the ListDefinitionCT complex type.
            this.VerifyListDefinitionCT(updateListResult.Results.ListProperties);

            // Verify the requirements of the UpdateListFieldResults complex type.
            if (updateListResult.Results != null)
            {
                this.VerifyUpdateListFieldResults(updateListResult.Results);
            }

            // Verify R2062
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2062,
                @"[UpdateListResponse]UpdateListResult: The results of the UpdateList request.");

            // Verify R2063
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2063,
                @"[UpdateListResponse]Results: The container element for the result categories.");

            // Verify R2064
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2064,
                @"[UpdateListResponse]NewFields: The container element for the results of any add  field requests. See section 2.2.4.14.");

            // Verify R2065
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2065,
                @"[UpdateListResponse]DeleteFields: The container element for the results of any delete field requests. See section 2.2.4.14.");

            // Verify R2066
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured.
            Site.CaptureRequirement(
                2066,
                @"[UpdateListResponse]UpdateFields: The container element for the results of any update field requests. See section 2.2.4.14.");
               
            //Verify requirement: MS-LISTSWS_R3010001
            if (Common.IsRequirementEnabled(3010001, this.Site))
            {
                if (!bool.Parse(updateListResult.Results.ListProperties.HasRelatedLists))
                {
                    Site.CaptureRequirement(
                    3010001,
                    @"[ListDefinitionCT.HasRelatedLists] Otherwise [if this list does not have any related lists] is ""False"".");
                }
            }
            
        }

        /// <summary>
        /// Verify the message syntax of UpdateListItems operation when the response is received 
        /// successfully.
        /// </summary>
        /// <param name="updateListItemsResult">The result of the operation.</param>
        /// <param name="updates">The updates parameter of the method.</param>
        private void VerifyUpdateListItemsOperation(
            UpdateListItemsResponseUpdateListItemsResult updateListItemsResult,
            UpdateListItemsUpdates updates)
        {
            // Verify R2067
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2067,
                @"[The schema of UpdateListItems is defined as:]"
                + @"<wsdl:operation name=""UpdateListItems"">"
                + @"    <wsdl:input message=""UpdateListItemsSoapIn"" />"
                + @"    <wsdl:output message=""UpdateListItemsSoapOut"" />"
                + @"</wsdl:operation>");

            // Verify R905
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                905,
                @"[In UpdateListItems operation] [If the protocol client sends an "
                + "UpdateListItemsSoapIn request message] the protocol server responds with an "
                + "UpdateListItemsSoapOut response message.");

            // Verify R2072
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2072,
                @"[UpdateListItemsSoapOut]The SOAP Body contains an UpdateListItemsResponse "
                + "element");

            // Verify R2111
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2111,
                @"[The schema of UpdateListItemsResponse is defined as:]  <s:element name=""UpdateListItemsResponse"">
                <s:complexType>
                  <s:sequence>
                    <s:element minOccurs=""0"" maxOccurs=""1"" name=""UpdateListItemsResult"">
                      <s:complexType mixed=""true"">
                        <s:sequence>
                          <s:element name=""Results"" >
                            <s:complexType>
                              <s:sequence>
                                <s:element name=""Result"" maxOccurs=""unbounded"">
                                  <s:complexType>
                                    <s:sequence>
                                      <s:element name=""ErrorCode"" type=""s:string"" />
                                              <s:any minOccurs=""0"" maxOccurs=""unbounded"" />
                                    </s:sequence>
                                  <s:attribute name=""ID"" type=""s:string"" />
                                <s:attribute name=""List"" type=""s:string""/>
                              <s:attribute name=""Version"" type=""s:string""/>
                            </s:complexType>
                          </s:element>
                        </s:sequence>
                      </s:complexType>
                    </s:element>
                  </s:sequence>
                </s:complexType>
              </s:element>
            </s:sequence>
          </s:complexType></s:element>");

            // Ensure the SOAP result is de-serialized successfully.
            Site.Assume.IsNotNull(
                updateListItemsResult,
                "The result of UpdateListItems operation must not be null.");

            // If the first returned ID attribute is the Method ID, followed by a comma, 
            // followed by the Method operation, then the following requirements can be 
            // captured.
            for (int i = 0; i < updates.Batch.Method.Length; i++)
            {
                bool isVerifyR962 = false;
                bool isVerifyR923 = false;
                string[] strID = updateListItemsResult.Results[i].ID.Split(',');
                if (strID[0] == updates.Batch.Method[i].ID.ToString() && strID[1] == updates.Batch.Method[i].Cmd.ToString())
                {
                    isVerifyR962 = true;
                    isVerifyR923 = true;
                }

                // Verify R962
                Site.CaptureRequirementIfIsTrue(
                isVerifyR962,
                962,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In "
                + "UpdateListItemsResult element] The ID attribute of the Method parameters "
                + "MUST correspond to the ID attribute of the Result element and the Result ID "
                + "is the Method ID, followed by a comma, followed by the Method operation.");

                // Verify R923
                Site.CaptureRequirementIfIsTrue(
                isVerifyR923,
                923,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates "
                + "element] [In Method element] If the Method ID attribute is unique, the protocol "
                + "server MUST use the method identification to match up the request made to the "
                + "protocol server with the protocol server response.");
                //
                if ((updateListItemsResult.Results[i].ID == null || updateListItemsResult.Results[i].ID.ToString() == "") && (updateListItemsResult.Results[i].ErrorCode != ""))
                {
                Site.CaptureRequirement(
                    2323001,
                    @"An empty ID element following the ErrorCode element is included, which is reserved for future use. ");
                }
            }

            // Verify R2115
            // The response have been received successfully, then the following requirement can be captured.
            // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
            Site.CaptureRequirement(
                2115,
                @"[UpdateListItemsResponse]The Result element MUST contain an ErrorCode element.");
        }

        /// <summary>
        /// Verify the message syntax of UpdateListItemsWithKnowledge operation when the 
        /// response is received successfully.
        /// </summary>
        /// <param name="updateListItemsWithKnowledgeResult">The result of the operation.</param>
        /// <param name="updates">The updates parameter of the method.</param>
        private void VerifyUpdateListItemsWithKnowledgeOperation(
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult updateListItemsWithKnowledgeResult,
            UpdateListItemsWithKnowledgeUpdates updates)
        {
            if (updateListItemsWithKnowledgeResult != null)
            {
                // Verify R1148
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1148,
                    @"[The schema of UpdateListItemsWithKnowledge is defined as:]"
                    + @"<wsdl:operation name=""UpdateListItemsWithKnowledge"">"
                    + @"    <wsdl:input message=""UpdateListItemsWithKnowledgeSoapIn"" />"
                    + @"    <wsdl:output message=""UpdateListItemsWithKnowledgeSoapOut"" />"
                    + @"</wsdl:operation>");

                // Verify R1150
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1150,
                    @"[In UpdateListItemsWithKnowledge]the protocol server responds with an "
                    + "UpdateListItemsWithKnowledgeSoapOut response message,");

                // Verify R1163
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1163,
                    @"[In UpdateListItemsWithKnowledgeSoapOut]The SOAP Body contains an "
                    + "UpdateListItemsWithKnowledgeResponse element.");

                // Verify R1175
                // The response have been received successfully, then the following requirement can be captured.
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    1175,
                    @"[The schema of UpdateListItemsWithKnowledgeresponse is defined as: ] <s:element name=""UpdateListItemsWithKnowledgeResponse"">
                  <s:complexType>
                    <s:sequence>
                      <s:element minOccurs=""0"" maxOccurs=""1"" name=""UpdateListItemsWithKnowledgeResult"">
                        <s:complexType mixed=""true"">
                          <s:sequence>
                            <s:element name=""Results"" >
                              <s:complexType>
                                <s:sequence>
                                  <s:element name=""Result"" maxOccurs=""unbounded"">
                                    <s:complexType>
                                      <s:sequence>
                                        <s:element name=""ErrorCode"" type=""s:string"" />
                                        <s:any minOccurs=""0"" maxOccurs=""unbounded""/>
                                      </s:sequence>
                                      <s:attribute name=""ID"" type=""s:string"" />
                                      <s:attribute name=""List"" type=""s:string"" />
                                      <s:attribute name=""Version"" type=""s:string"" />
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

                // If the first returned ID attribute is the Method ID, followed by a comma, 
                // followed by the Method operation, then the following requirements can be 
                // captured.
                for (int i = 0; i < updates.Batch.Method.Length; i++)
                {
                    bool isVerifyR2317 = false;
                    bool isVerifyR2350 = false;
                    string[] strID = updateListItemsWithKnowledgeResult.Results[i].ID.Split(',');
                    if (strID[0] == updates.Batch.Method[i].ID.ToString() && strID[1] == updates.Batch.Method[i].Cmd.ToString())
                    {
                        isVerifyR2317 = true;
                        isVerifyR2350 = true;
                    }

                    // Verify R2317
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2317,
                        2317,
                        @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                        + "[In UpdateListItemsWithKnowledgeResult element]The ID attribute of the Method "
                        + "parameters MUST correspond to the ID attribute of the Result element and the "
                        + "Result ID is the Method ID, followed by a comma, followed by the Method operation.");

                    // Verify R2317
                    Site.CaptureRequirementIfIsTrue(
                    isVerifyR2350,
                    2350,
                    @"[In UpdateListItemsWithKnowledge operation] [In "
                    + "UpdateListItemsWithKnowledge element] [In updates element] [In Method "
                    + "element] If the Method ID attribute is unique, the protocol server MUST use "
                    + "the method identification to match up the request made to the protocol server "
                    + "with the protocol server response.");
                }

                // Verify R2322
                // If the response is not received and parsed successfully, the test case will fail before this requirement is captured 
                Site.CaptureRequirement(
                    2322,
                    @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                    + "[In UpdateListItemsWithKnowledgeResult element]The Result element MUST contain "
                    + "an ErrorCode element.");
            }
        }
        #endregion
    }
}