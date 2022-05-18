namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class includes the capture methods.
    /// </summary>
    public partial class MS_WWSPAdapter
    {
        #region Variable

        /// <summary>
        /// The XML Schema validation result returned by WWSP Web service operations.
        /// </summary>
        private static bool isSuccess = false;

        #endregion Variable

        #region CaptureRequirements

        #region CaptureRelatedRequirements_GetTemplatesForItem

        /// <summary>
        /// Capture XML schema related requirements of GetTemplatesForItem operation.
        /// </summary>
        private void CaptureXMLSchemaGetTemplatesForItem()
        {
            // The schema of GetTemplatesForItem operation has been validated by full WSDL.
            // If it returns success, the schema of GetTemplatesForItem operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R149
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                149,
                @"[In GetTemplatesForItem][The schema of the GetTemplatesForItem is:] <wsdl:operation name=""GetTemplatesForItem"">
    <wsdl:input message=""GetTemplatesForItemSoapIn"" />
    <wsdl:output message=""GetTemplatesForItemSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R171
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                171,
                @"[In GetTemplatesForItemResponse][The schema of the GetTemplatesForItemResponse is :] <s:element name=""GetTemplatesForItemResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetTemplatesForItemResult"" minOccurs=""0"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element ref=""tns:TemplateData"" minOccurs=""1"" maxOccurs=""1""/>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture soap information related requirements of GetTemplatesForItem operation.
        /// </summary>
        private void CaptureSoapInfoGetTemplatesForItem()
        {
            // Check whether SOAP body contains a GetTemplatesForItemResponse element.
            bool isVerifyR159 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "GetTemplatesForItemResponse");

            // Verify MS-WWSP requirement: MS-WWSP_R159
            Site.CaptureRequirementIfIsTrue(
                isVerifyR159,
                159,
                @"[In GetTemplatesForItemSoapOut] The SOAP body contains a GetTemplatesForItemResponse element.");

            // Check whether GetTemplatesForItemResult contains a TemplateData element.
            bool isVerifyR172 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "TemplateData");

            // Verify MS-WWSP requirement: MS-WWSP_R172
            Site.CaptureRequirementIfIsTrue(
                isVerifyR172,
                172,
                @"[In GetTemplatesForItemResponse] GetTemplatesForItemResult: Contains a TemplateData element as specified in section 2.2.3.1 that specifies a set of workflow associations.");

            // Check whether GetTemplatesForItemResponse is sent with GetTemplatesForItemSoapOut.
            bool isVerifyR169 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "GetTemplatesForItemResponse");

            // Verify MS-WWSP requirement: MS-WWSP_R169
            Site.CaptureRequirementIfIsTrue(
                isVerifyR169,
                169,
                "[In GetTemplatesForItemResponse] This element is sent with a GetTemplatesForItemSoapOut message.");
        }

        #endregion CaptureRelatedRequirements_GetTemplatesForItem

        #region CaptureRelatedRequirements_GetToDosForItem

        /// <summary>
        /// Capture XML schema related requirements of GetToDosForItem operation.
        /// </summary>
        private void CaptureXMLSchemaGetToDosForItem()
        {
            // The schema of GetToDosForItem operation has been validated by full WSDL.
            // If it returns success, the schema of GetToDosForItem operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R175
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                175,
                @"[In GetToDosForItem][The schema of the GetToDosForItem is:] <wsdl:operation name=""GetToDosForItem"">
    <wsdl:input message=""GetToDosForItemSoapIn"" />
    <wsdl:output message=""GetToDosForItemSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R193
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                193,
               @"[In GetToDosForItemResponse][The schema of the GetToDosForItemResponse is:] <s:element name=""GetToDosForItemResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetToDosForItemResult"" minOccurs=""0"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element ref=""tns:ToDoData"" minOccurs=""1"" maxOccurs=""1""/>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture soap information related requirements of GetToDosForItem operation.
        /// </summary>
        private void CaptureSoapInfoGetToDosForItem()
        {
            // Verify MS-WWSP requirement: MS-WWSP_R185
            bool isVerifyR185 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "GetToDosForItemResponse");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-WWSP_R189. Whether SOAP body contains a GetToDosForItemResponse element: {0}", isVerifyR185);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR185,
                185,
                "[In GetToDosForItemSoapOut]The SOAP body contains a GetToDosForItemResponse element.");

            // Verify MS-WWSP requirement: MS-WWSP_R194
            bool isVerifyR194 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "ToDoData");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-WWSP_R195. Whether GetToDosForItemResult contains a ToDoData element: {0}", isVerifyR194);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR194,
                194,
                "[In GetToDosForItemResponse]GetToDosForItemResult: Contains a ToDoData element as specified in section 2.2.3.2 containing data about a set of workflow tasks.");
        }

        #endregion CaptureRelatedRequirements_GetToDosForItem

        #region CaptureRelatedRequirements_GetWorkflowDataForItem

        /// <summary>
        /// Capture XML schema related requirements of GetWorkflowDataForItem operation.
        /// </summary>
        private void CaptureXMLSchemaGetWorkflowDataForItem()
        {
            // The schema of GetWorkflowDataForItem operation has been validated by full WSDL.
            // If it returns success, the schema of GetWorkflowDataForItem operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R197
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                197,
                @"[In GetWorkflowDataForItem][The schema of the GetWorkflowDataForItem is:] <wsdl:operation name=""GetWorkflowDataForItem"">
    <wsdl:input message=""GetWorkflowDataForItemSoapIn"" />
    <wsdl:output message=""GetWorkflowDataForItemSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R217
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                217,
                @"[In GetWorkflowDataForItemResponse][The schema of the GetWorkflowDataForItemResponse is:] <s:element name=""GetWorkflowDataForItemResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetWorkflowDataForItemResult"" minOccurs=""1"" maxOccurs=""1"">
        <s:complexType>
          <s:sequence>
            <s:element name=""WorkflowData"" minOccurs=""1"" maxOccurs=""1"">
              <s:complexType>
                <s:sequence>
                  <s:element ref=""tns:ToDoData"" minOccurs=""1"" maxOccurs=""1"" />
                  <s:element ref=""tns:TemplateData"" minOccurs=""1"" maxOccurs=""1"" />
                  <s:element name=""ActiveWorkflowsData"" minOccurs=""1"" maxOccurs=""1"" >
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""Workflows"" minOccurs=""1"" maxOccurs=""1"" >
                          <s:complexType>
                            <s:sequence>
                              <s:element name=""Workflow"" minOccurs=""0"" maxOccurs=""unbounded"">
                                <s:complexType>
                                  <s:attribute name=""StatusPageUrl"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Id"" type=""s1:guid"" use=""required"" />
                                  <s:attribute name=""TemplateId"" type=""s1:guid"" use=""required"" />
                                  <s:attribute name=""ListId"" type=""s1:guid"" use=""required""/>
                                  <s:attribute name=""SiteId"" type=""s1:guid"" use=""required"" />
                                  <s:attribute name=""WebId"" type=""s1:guid"" use=""required""/>
                                  <s:attribute name=""ItemId"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""ItemGUID"" type=""s1:guid"" use=""required""/>
                                  <s:attribute name=""TaskListId"" type=""s1:guid"" use=""required""/>
                                  <s:attribute name=""AdminTaskListId"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Author"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""Modified"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Created"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""StatusVersion"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""Status1"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""Status2"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status3"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status4"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status5"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status6"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status7"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status8"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status9"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Status10"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""TextStatus1"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""TextStatus2"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""TextStatus3"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""TextStatus4"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""TextStatus5"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""Modifications"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""ActivityDetails"" type=""s:string"" use=""required"" />
                                  <s:attribute name=""CorrelationId"" type="" s:string"" use=""required"" />
                                  <s:attribute name=""InstanceData"" type=""s:string"" use=""required""/>
                                  <s:attribute name=""InstanceDataSize"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""InternalState"" type=""s:int"" use=""required""/>
                                  <s:attribute name=""ProcessingId"" type=""s:string"" use=""required""/>
                                </s:complexType>
                              </s:element>
                            </s:sequence>
                          </s:complexType>
                        </s:element>
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                  <s:element name=""DefaultWorkflows"" minOccurs=""1"" maxOccurs=""1"" >
                    <s:complexType>
                      <s:sequence>
                        <s:element name=""DefaultWorkflow"" minOccurs=""0"" maxOccurs=""1"" >
                          <s:complexType>
                            <s:attribute name=""Event"" type=""s:string"" use=""required""/>
                            <s:attribute name=""TemplateId"" type=""s1:guid"" use=""required""/>
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
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify R218
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                218,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.ToDoData: Specifies a set of workflow tasks as defined in section 2.2.3.2.");

            // Verify R219
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                219,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.TemplateData: Specifies a set of workflow associations as defined in section 2.2.3.1.");
        }

        /// <summary>
        /// Capture soap information related requirements of GetWorkflowDataForItem operation.
        /// </summary>
        /// <param name="responseOfGetWorkflowDataForItem">A parameter represents the response Of GetWorkflowDataForItem operation.</param>
        private void CaptureGetWorkflowDataForItem(GetWorkflowDataForItemResponseGetWorkflowDataForItemResult responseOfGetWorkflowDataForItem)
        {  
            // Verify MS-WWSP requirement: MS-WWSP_R207
            bool isVerifyR207 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "GetWorkflowDataForItemResponse");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR207,
                207,
                "[In GetWorkflowDataForItemSoapOut] The SOAP body contains a GetWorkflowDataForItemResponse element.");

            this.VerifyWorkflowItemInGetWorkflowDataForItemResponse(responseOfGetWorkflowDataForItem);
        }

        #endregion CaptureRelatedRequirements_GetWorkflowDataForItem

        #region CaptureRelatedRequirements_StartWorkflow

        /// <summary>
        /// Capture XML schema related requirements of StartWorkflow operation.
        /// </summary>
        private void CaptureXMLSchemaStartWorkflow()
        {
            // The schema of StartWorkflow operation has been validated by full WSDL.
            // If it returns success, the schema of StartWorkflow operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R326
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                326,
                @"[In StartWorkflow][The schema of the StartWorkflow is:] <wsdl:operation name=""StartWorkflow"">
    <wsdl:input message=""StartWorkflowSoapIn"" />
    <wsdl:output message=""StartWorkflowSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R352
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                352,
                @"[In StartWorkflowResponse][The schema of the StartWorkflowResponse is:] 
  <s:element name=""StartWorkflowResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""StartWorkflowResult"" minOccurs=""1""/>
    </s:sequence>	
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture soap information related requirements of StartWorkflow operation.
        /// </summary>
        private void CaptureSoapInfoStartWorkflow()
        {
            // Verify MS-WWSP requirement: MS-WWSP_R336
            bool isVerifyR336 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "StartWorkflowResponse");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR336,
                336,
                @"[In StartWorkflowSoapOut] The SOAP body contains a StartWorkflowResponse element.");
        }

        #endregion CaptureRelatedRequirements_StartWorkflow

        #region CaptureRelatedRequirements_GetWorkflowTaskData

        /// <summary>
        /// Capture XML schema related requirements of GetWorkflowTaskData operation.
        /// </summary>
        private void CaptureXMLSchemaGetWorkflowTaskData()
        {
            // The schema of GetWorkflowTaskData operation has been validated by full WSDL.
            // If it returns success, the schema of GetWorkflowTaskData operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R294
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                294,
                @"[In GetWorkflowTaskData][The schema of the GetWorkflowTaskData is: ] <wsdl:operation name=""GetWorkflowTaskData"">
    <wsdl:input message=""GetWorkflowTaskDataSoapIn"" />
    <wsdl:output message=""GetWorkflowTaskDataSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R320
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                320,
                @"[In GetWorkflowTaskDataResponse][The schema of the GetWorkflowTaskDataResponse is:] <s:element name=""GetWorkflowTaskDataResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""GetWorkflowTaskDataResult"" >
        <s:complexType>
         <s:sequence>
          <s:any minOccurs=""0"" maxOccurs=""unbounded"" />
         </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Verify MS-PRSTFR requirement: MS-PRSTFR_R179
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                "MS-PRSTFR",
                179,
                @"[The complexType of rowset namespace] The following XML schema defines a subset of rowset namespace which is relevant in this context:
<xs:schema xmlns=""urn:schemas-microsoft-com:rowset"" 
            targetNamespace=""urn:schemas-microsoft-com:rowset""
            attributeFormDefault=""unqualified""
            elementFormDefault=""qualified"" 
            xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:attribute name=""name"" type=""xs:string"" />
  <xs:attribute name=""number"" type=""xs:int"" />
  <xs:attribute name=""precision"" type=""xs:unsignedByte"" />
  <xs:attribute name=""scale"" type=""xs:unsignedByte"" />
  <xs:attribute name=""CommandTimeout"" type=""xs:int""/>
  <xs:complexType name=""data"">
    <xs:sequence minOccurs=""0"" maxOccurs=""unbounded"">
      <xs:any minOccurs=""0"" maxOccurs=""unbounded"" namespace=""##any""
              processContents=""lax"" />
    </xs:sequence>
  </xs:complexType>
</xs:schema>");

            // The schema of GetWorkflowTaskData operation has been validated by full WSDL.
            // If it returns success, the schema of GetWorkflowTaskData operation is valid, and the type validation for each attribute is valid.
            // When validating full WSDL is correct which includes the type validation for each attribute, we can capture the following requirements. 
            // Verify MS-PRSTFR requirement: MS-PRSTFR_R185
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                "MS-PRSTFR",
                185,
                @"[The complexType of rowset namespace] precision: Precision is the number of digits in a floating point number.");

            // Verify MS-PRSTFR requirement: MS-PRSTFR_R187
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                "MS-PRSTFR",
                187,
                @"[The complexType of rowset namespace] scale: Scale is the number of digits to the right of the decimal point.");

            // Verify R321
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                321,
                @"[In GetWorkflowTaskDataResponse] GetWorkflowTaskDataResult: GetWorkflowTaskDataResult element contains an array of z:row elements, where the z is equal to #RowsetSchema in the ActiveX Data Objects (ADO) XML Persistence format (see [MS-PRSTFR]).");
        }

        /// <summary>
        /// Capture soap information related requirements of GetWorkflowTaskData operation.
        /// </summary>
        private void CaptureSoapInfoGetWorkflowTaskData()
        {
            // Verify MS-WWSP requirement: MS-WWSP_R304
            bool isVerifyR304 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "GetWorkflowTaskDataResponse");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR304,
                304,
                "[In GetWorkflowTaskDataSoapOut] The SOAP body contains a GetWorkflowTaskDataResponse element.");
        }

        #endregion CaptureRelatedRequirements_GetWorkflowTaskData

        #region CaptureRelatedRequirements_AlterToDo

        /// <summary>
        /// Capture XML schema related requirements of AlterToDo operation.
        /// </summary>
        private void CaptureXMLSchemaAlterToDo()
        {
            // The schema of AlterToDo operation has been validated by full WSDL.
            // If it returns success, the schema of AlterToDo operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R79
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                79,
                @"[In AlterToDo][The AlterToDo schema is:] <wsdl:operation name=""AlterToDo"">
    <wsdl:input message=""AlterToDoSoapIn"" />
    <wsdl:output message=""AlterToDoSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R104
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                104,
               @"[In AlterToDoResponse][The schema of the AlterToDoResponse is:] <s:element name=""AlterToDoResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""AlterToDoResult"" minOccurs=""1"" maxOccurs=""1"" >
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""fSuccess"" type=""s:int"" minOccurs=""1"" maxOccurs=""1"" />
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture soap information related requirements of AlterToDo operation.
        /// </summary>
        private void CaptureSoapInfoAlterToDo()
        {
            // Check whether SOAP body contains AlterToDoResponse element.
            bool isVerifyR89 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "AlterToDoResponse");

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-WWSP_R89, whether the SOAP body contains an AlterToDoResponse element:{0}", isVerifyR89);

            // Verify MS-WWSP requirement: MS-WWSP_R89
            Site.CaptureRequirementIfIsTrue(
                isVerifyR89,
                89,
                "[In AlterToDoSoapOut] The SOAP body contains an AlterToDoResponse element.");
        }

        #endregion CaptureRelatedRequirements_AlterToDo

        #region CaptureRelatedRequirements_ClaimReleaseTask

        /// <summary>
        /// Capture XML schema related requirements of ClaimReleaseTask operation.
        /// </summary>
        private void CaptureXMLSchemaClaimReleaseTask()
        {
            // The schema of ClaimReleaseTask operation has been validated by full WSDL.
            // If it returns success, the schema of ClaimReleaseTask operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Validated that full WSDL(using SchemaValidation.cs) is correct or not: {0}", isSuccess);

            // Verify MS-WWSP requirement: MS-WWSP_R112
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                112,
                @"[In ClaimReleaseTask][The schema of the ClaimReleaseTask is:] <wsdl:operation name=""ClaimReleaseTask"">
    <wsdl:input message=""ClaimReleaseTaskSoapIn"" />
    <wsdl:output message=""ClaimReleaseTaskSoapOut"" />
</wsdl:operation>");

            // Verify MS-WWSP requirement: MS-WWSP_R141
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                141,
               @"[In ClaimReleaseTaskResponse][The schema of the ClaimReleaseTaskResponse is:] <s:element name=""ClaimReleaseTaskResponse"">
  <s:complexType>
    <s:sequence>
      <s:element name=""ClaimReleaseTaskResult"" minOccurs=""1"">
        <s:complexType mixed=""true"">
          <s:sequence>
            <s:element name=""TaskData"" minOccurs=""1"" maxOccurs=""1"">
              <s:complexType>
                <s:attribute name=""AssignedTo"" type=""s:string"" use=""required"">
                <s:attribute name=""TaskGroup"" type=""s:string"" use=""required"">
                <s:attribute name=""ItemId"" type=""s:int"" use=""required"">
                <s:attribute name=""ListId"" type=""s1:guid"" use=""required"">
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");
        }

        /// <summary>
        /// Capture soap information related requirements of ClaimReleaseTask operation.
        /// </summary>
        private void CaptureSoapInfoClaimReleaseTask()
        {
            // Check whether the SOAP body contains a ClaimReleaseTaskResponse element.
            bool isVerifyR122 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "ClaimReleaseTaskResponse");

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-WWSP_R122.Whether the SOAP body contains a ClaimReleaseTaskResponse element:{0}", isVerifyR122);

            // Verify MS-WWSP requirement: MS-WWSP_R122
            Site.CaptureRequirementIfIsTrue(
                isVerifyR122,
                122,
                "[In ClaimReleaseTaskSoapOut] The SOAP body contains a ClaimReleaseTaskResponse element.");

            // Check whether ClaimReleaseTaskResponse is sent with ClaimReleaseTaskSoapOut.
            bool isVerifyR138 = AdapterHelper.HasElement(SchemaValidation.LastRawResponseXml, "ClaimReleaseTaskResponse");

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-WWSP_R138.Whether ClaimReleaseTaskResponse is sent with ClaimReleaseTaskSoapOut:{0}", isVerifyR138);

            // Verify MS-WWSP requirement: MS-WWSP_R138
            Site.CaptureRequirementIfIsTrue(
                isVerifyR138,
                138,
                "[In ClaimReleaseTaskResponse] This element[ClaimReleaseTaskResponse] is sent with ClaimReleaseTaskSoapOut.");
        }

        #endregion CaptureRelatedRequirements_ClaimReleaseTask

        #region CaptureRelatedRequirements_TemplateData

        /// <summary>
        /// This method is used to capture the TemplateData related requirements.
        /// </summary>
        /// <param name="templateData">A parameter represents the template data get from SUT.</param>
        private void CaptureTemplateDataRelatedRequirements(TemplateData templateData)
        {
            // The schema has been validated by full WSDL.
            // If it returns success, the schema is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;
            
            // Verify MS-WWSP requirement: MS-WWSP_R34
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                34,
               @"[In TemplateData][The TemplateData schema is:] <s:element name=""TemplateData"" >
  <s:complexType>
    <s:sequence>
      <s:element name=""Web"" minOccurs=""1"" maxOccurs=""1"" >
        <s:complexType>
          <s:attribute name=""Title"" type=""s:string"" use=""required"" />
          <s:attribute name=""Url"" type=""s:string"" use=""required"" />
        </s:complexType>
      </s:element>
      <s:element name=""List"" minOccurs=""1"" maxOccurs=""1"" >
        <s:complexType>
          <s:attribute name=""Title"" type=""s:string"" use=""required"" />
          <s:attribute name=""Url"" type=""s:string"" use=""required"" />
        </s:complexType>
      </s:element>
      <s:element name=""WorkflowTemplates"" >
        <s:complexType>
          <s:sequence>
            <s:element name=""WorkflowTemplate"" minOccurs=""0"" maxOccurs=""unbounded"">
              <s:complexType>
                <s:sequence>
                  <s:element name=""WorkflowTemplateIdSet"" minOccurs=""1"" maxOccurs=""1"">
                    <s:complexType>
                      <s:attribute name=""TemplateId"" type=""s1:guid"" use=""required"" />
                      <s:attribute name=""BaseId"" type=""s1:guid"" use=""required"" />
                    </s:complexType>
                  </s:element>
                  <s:element name=""AssociationData"" minOccurs=""0"" maxOccurs=""1"" >
                    <s:complexType>
                      <s:sequence>
                        <s:any/>
                      </s:sequence>
                    </s:complexType>
                  </s:element>
                  <s:element name=""Metadata"" minOccurs=""1"" maxOccurs=""1"">
                    <s:complexType>
                      
                        <s:all>
                          <s:element name=""InitiationCategories"" minOccurs=""0"" maxOccurs=""1"">
                            <s:complexType>
                              <s:sequence>
                                <s:any/>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                          <s:element name=""Instantiation_FormURN"" minOccurs=""0"" maxOccurs=""1"">
                            <s:complexType>
                              <s:sequence>
                                <s:any/>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                          <s:element name=""Instantiation_FormURI"" minOccurs=""0"" maxOccurs=""1"">
                            <s:complexType>
                              <s:sequence>
                                <s:any/>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                          <s:element name=""AssignmentStagesName"" minOccurs=""0"" maxOccurs=""1"">
                            <s:complexType>
                              <s:sequence>
                                <s:any/>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                          <s:element name=""SigClientSettings"" minOccurs=""0"" maxOccurs=""1"">
                            <s:complexType>
                              <s:sequence>
                                <s:any/>
                              </s:sequence>
                            </s:complexType>
                          </s:element>
                        </s:all>
                      
                    </s:complexType>
                  </s:element>
                </s:sequence>
                <s:attribute name=""Name"" type=""s:string"" use=""required"" />
                <s:attribute name=""Description"" type=""s:string"" use=""required"" />
                <s:attribute name=""InstantiationUrl"" type=""s:string"" />
              </s:complexType>
            </s:element>
          </s:sequence>
        </s:complexType>
      </s:element>
    </s:sequence>
  </s:complexType>
</s:element>");

            // Capture R421
           if (Common.IsRequirementEnabled(421, this.Site))
           {
               var instantiationFormURIElements = from workflowTemplateItem in templateData.WorkflowTemplates
                                  where workflowTemplateItem.Metadata.Instantiation_FormURI != null
                                       && !string.IsNullOrEmpty(workflowTemplateItem.Metadata.Instantiation_FormURI.InnerText)
                                  select workflowTemplateItem.Metadata.Instantiation_FormURI;
               
               int numberOfexistingElements = instantiationFormURIElements.Count();
               this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"The actual value: instantiationFormURI Elements' number[{0}] for requirement #R421",
                                numberOfexistingElements);
                             
               this.Site.CaptureRequirementIfAreNotEqual(
                                            0,
                                            numberOfexistingElements,
                                            421,
                                            @"[In Appendix B: Product Behavior]Implementation does support this[Instantiation_FormURI] element.(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.)");
           }

            // Get the association data which is included in current used workflow association item.
            string expectedworkflowTemplateName = Common.GetConfigurationPropertyValue("WorkflowAssociationName", Site);
            XmlNode currentAssociationData = AdapterHelper.GetAssociationDataFromTemplateItem(expectedworkflowTemplateName, templateData);
           
            // Verify the association data relate requirements.
            this.CaptureWorkflowAssociation(currentAssociationData);
        }

        #endregion CaptureRelatedRequirements_TemplateData

        #region CaptureRelatedRequirements_ToDoData

        /// <summary>
        /// This method is used to toDoData element related requirements.
        /// </summary>
        /// <param name="tododata">A parameter represents the ToDoData which is contained in response.</param>
        private void CaptureToDoDataRelatedRequirements(ToDoData tododata)
        {
            if (null == tododata)
            {
                this.Site.Assert.Fail("The ToDoData should be contained in response.");
            }

            // If it returns success, the schema of GetWorkflowTaskData operation is valid, capture related requirements.
            isSuccess = SchemaValidation.ValidationResult == ValidationResult.Success;

            // Capture R55 and R56
            this.Site.CaptureRequirementIfIsTrue(
                                         isSuccess,
                                         55,
                                         @"[In ToDoData] The ToDoData element specifies a set of workflow tasks for a protocol client as follows:");

            this.Site.CaptureRequirementIfIsTrue(
                                        isSuccess,
                                        56,
                                        @"[In ToDoData][The ToDoData schema is:]  <s:element name=""ToDoData"" >
  <s:complexType>
    <s:sequence>
      <s:element name=""xml"" minOccurs=""0"" maxOccurs=""1"" >
        <s:complexType>
	      <s:sequence>
	        <s:element ref=""rs: data"" maxOccurs=""1"" />
          </ s:sequence >
        </ s:complexType >
      </ s:element >
    </ s:sequence>
  </s:complexType>
</s:element>");

            // Capture #R406, all schema definitions in [MS-PRSTFR] section 2.4 are validate in schema validation process.
            this.Site.CaptureRequirementIfIsTrue(
                                        isSuccess,
                                        406,
                                        @"[In ToDoData] xml: A set of Rowsets as specified in [MS-PRSTFR] section 2.4.");
        }

        #endregion CaptureRelatedRequirements_ToDoData

        /// <summary>
        /// This method is used to capture the transport type and soap version related requirements.
        /// </summary>
        private void CaptureTranstportAndSOAPRequirements()
        {
            TransportProtocol currentTransportValue = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (currentTransportValue == TransportProtocol.HTTP)
            {   
                // If current test suite use HTTP as the low level transport protocol, then capture R3.
                // Verify MS-WWSP requirement: MS-WWSP_R3
                this.Site.CaptureRequirement(
                                        3,
                                        @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }
            else
            {   
                if (Common.IsRequirementEnabled(5, this.Site))
                {
                    // If current test suite use HTTPS as the low level transport protocol, then capture R5.
                    // Verify MS-WWSP requirement: MS-WWSP_R5
                    this.Site.CaptureRequirement(
                                       5,
                                       @"[In Appendix B: Product Behavior] Protocol servers does additionally support SOAP over HTTPS for securing communication with clients.(Windows® SharePoint® Services 3.0 and above follow this behavior.)");
                }
            }

            SoapVersion currentSoapValue = Common.GetConfigurationPropertyValue<SoapVersion>("SOAPVersion", this.Site);
            if (currentSoapValue == SoapVersion.SOAP11)
            {
                // If current test suite use SOAP1.1 as the soap version, then capture R6.
                // Verify MS-WWSP requirement: MS-WWSP_R6
                this.Site.CaptureRequirement(
                                       6,
                                       @"[In Transport] Protocol messages MUST be formatted as specified in [SOAP1.1] section 4[, or [SOAP1.2-1/2007] section 5].");
            }
            else
            {
                // If current test suite use SOAP1.2 as the soap version, then capture R7.
                // Verify MS-WWSP requirement: MS-WWSP_R7
                this.Site.CaptureRequirement(
                                       7,
                                       @"[In Transport] Protocol messages MUST be formatted as specified in [[SOAP1.1] section 4, or] [SOAP1.2-1/2007] section 5.");
            }
        }

        /// <summary>
        /// This method is used to capture the HTTP protocol status type and soap version related requirements.
        /// </summary>
        private void CaptureHTTPStatusAndSoapFaultRequirements()
        {   
            // If there are any SOAP fault return, then capture R368.
            // Verify MS-WWSP requirement: MS-WWSP_R368 
            this.Site.CaptureRequirement(
                                368,
                                @"[In Transport] Protocol server faults MUST be returned either using HTTP Status Codes as specified in [RFC2616] section 10 or using SOAP faults as specified in [SOAP1.1] section 4.4, or [SOAP1.2-1/2007] section 5.4, SOAP Fault.");
        }

        /// <summary>
        /// This method is used to workFlowAssociationData element related requirements.
        /// </summary>
        /// <param name="workFlowAssociationData">A parameter represents the data of the association</param>
        private void CaptureWorkflowAssociation(XmlNode workFlowAssociationData)
        {
            // Verify R395
            this.Site.CaptureRequirementIfAreEqual<int>(
                0,
                workFlowAssociationData.Attributes.Count,
                395,
                @"[In TemplateData] This element[WorkflowTemplates.WorkflowTemplate.AssociationData] MUST contain no attributes.");

            // Verify R397
            if (Common.IsRequirementEnabled(397, this.Site))
            {
                string workFlowAssociationInnerValue = workFlowAssociationData.InnerText;
                this.Site.Log.Add(
                                 LogEntryKind.Debug,
                                 "The actual value: workFlowAssociationInnerValue[{0}]  for requirement #R397.",
                                 string.IsNullOrEmpty(workFlowAssociationInnerValue) ? "Null or Empty" : workFlowAssociationInnerValue);

                this.Site.CaptureRequirementIfIsFalse(
                                        string.IsNullOrEmpty(workFlowAssociationInnerValue),
                                        397,
                                        @"[In Appendix B: Product Behavior] Implementation does [WorkflowTemplates.WorkflowTemplate.AssociationData]  contain child elements or inner text. (Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }
         
            // Validate the schema of the workflow association data.
            AdapterHelper.VerifyWorkflowAssociationSchema(workFlowAssociationData);
            bool isverifyworkflowAssociationDataSchema = Common.GetConfigurationPropertyValue<bool>("ValidateWorkFlowAssociation", this.Site);

            if (isverifyworkflowAssociationDataSchema)
            {
                // If verify the workflow association data's schema, means expected workflow association data is return.
                // And this check is optional, it is determined by "ValidateWorkFlowAssociation" property in configuration file, and MS-WWSP_R388 can be directly covered.
                this.Site.CaptureRequirement(
                            388,
                            @"[In TemplateData] WorkflowTemplates.WorkflowTemplate: A workflow association.");
            }
        }

        /// <summary>
        /// This method is used to workFlow element related requirements.
        /// </summary>
        /// <param name="responseOfGetWorkflowDataForItem">A parameter represents the response Of GetWorkflowDataForItem operation.</param>
        private void VerifyWorkflowItemInGetWorkflowDataForItemResponse(GetWorkflowDataForItemResponseGetWorkflowDataForItemResult responseOfGetWorkflowDataForItem)
        {
            if (null == responseOfGetWorkflowDataForItem.WorkflowData.ActiveWorkflowsData.Workflows)
            {
                return;
            }

            #region Capture Workflows related requirement

            XmlNode workflows = AdapterHelper.GetNodeFromXML("Workflows", SchemaValidation.LastRawResponseXml);

            if (null == workflows.ChildNodes || 0 == workflows.ChildNodes.Count)
            {
                return;
            }

            List<string> invalidStatusPageUrlValues = new List<string>();
            List<string> invalidInternalStateValues = new List<string>();
            List<string> itemsWithOutActivityDetailsAttribute = new List<string>();
            List<string> itemsWithActivityDetailsAttribute = new List<string>();
            List<string> itemsWithOutCorrelationIdAttribute = new List<string>();
            List<string> itemsWithCorrelationIdAttribute = new List<string>();
            List<string> invalidIdValues = new List<string>();

            foreach (XmlNode workflowitem in workflows.ChildNodes)
            {
                #region Verify StatusPageUrl attribute.

                // Get the value of StatusPageUrl.
                string statusPageUrl = workflowitem.Attributes["StatusPageUrl"].InnerText;
                if (!string.IsNullOrEmpty(statusPageUrl))
                {
                    // Fully qualified URL means a URL that includes a protocol transport scheme name("http" or "https") and a host name.
                    // Absolute means URL is from the very start, e.g.http://abc.com/a/b/c.htm.
                    Uri statusPageUrlUriInstance;
                    bool isabsoluteAndFullyqualifiedURL = false;
                    if (Uri.TryCreate(statusPageUrl, UriKind.Absolute, out statusPageUrlUriInstance))
                    {
                        if (statusPageUrlUriInstance.HostNameType == UriHostNameType.Dns)
                        {
                            string currentsutComputerName = Common.GetConfigurationPropertyValue("SUTComputerName", this.Site);
                            isabsoluteAndFullyqualifiedURL = statusPageUrlUriInstance.Host.Equals(currentsutComputerName, StringComparison.OrdinalIgnoreCase);
                        }
                        else
                        {
                            isabsoluteAndFullyqualifiedURL = true;
                        }
                    }

                    if (!isabsoluteAndFullyqualifiedURL)
                    {
                        invalidStatusPageUrlValues.Add(statusPageUrl);
                    }
                }
                #endregion verify StatusPageUrl attributes.

                #region Verify InternalState attribute.

                // Get the value of InternalState.
                string internalStateValue = workflowitem.Attributes["InternalState"].InnerText;
                long internalState;
                bool isvalidinternalStateValue = false;
                if (long.TryParse(internalStateValue, out internalState))
                {
                    // The values in the array are the possible bitmasks for combining the value of internalState.
                    long[] bitMasks = new long[] { 0x00000001, 0x00000002, 0x00000004, 0x00000008, 0x00000040, 0x00000080, 0x00000100, 0x00000400, 0x00000800, 0x00001000 };

                    // Check whether the value of internalState is zero or more combination of the bitmasks.
                    isvalidinternalStateValue = AdapterHelper.IsValueValid(internalState, bitMasks);
                }

                if (!isvalidinternalStateValue)
                {
                    invalidInternalStateValues.Add(internalStateValue);
                }

                #endregion Verify InternalState attribute.

                #region Verify ActivityDetails attribute

                var withActivityDetails = from XmlAttribute attributeItem in workflowitem.Attributes
                                            where attributeItem.Name.Equals("ActivityDetails", StringComparison.OrdinalIgnoreCase)
                                            select attributeItem;

                if (withActivityDetails.Count() == 0)
                {
                    string workflowitemId = workflowitem.Attributes["Id"].Value;
                    itemsWithOutActivityDetailsAttribute.Add(workflowitemId);
                }    
                else if (withActivityDetails.Count() == 1)
                {   
                    string workflowitemId = workflowitem.Attributes["Id"].Value;
                    itemsWithActivityDetailsAttribute.Add(workflowitemId);
                }
                else
                {
                    this.Site.Assert.Fail(@"The work flow item have un-expected number [{0}] of [ActivityDetails] attribute.", withActivityDetails.Count());
                }

                #endregion Verify ActivityDetails attribute

                #region Verify CorrelationId Attribute

                var withCorrelationId = from XmlAttribute attributeItem in workflowitem.Attributes
                                               where attributeItem.Name.Equals("CorrelationId", StringComparison.OrdinalIgnoreCase)
                                               select attributeItem;

                if (withCorrelationId.Count() == 0)
                {
                    string workflowitemId = workflowitem.Attributes["Id"].Value;
                    itemsWithOutCorrelationIdAttribute.Add(workflowitemId);
                }
                else if (withCorrelationId.Count() == 1)
                {
                    string workflowitemId = workflowitem.Attributes["Id"].Value;
                    itemsWithCorrelationIdAttribute.Add(workflowitemId);
                }
                else
                {
                    this.Site.Assert.Fail(@"The work flow item have un-expected number [{0}] of [CorrelationId] attribute.", withCorrelationId.Count());
                }

                #endregion Verify CorrelationId Attribute

                #region Verify ID attribute

                string idattributeValue = workflowitem.Attributes["Id"].InnerText;
                if (string.IsNullOrEmpty(idattributeValue))
                {
                    invalidIdValues.Add("EmptyOrNull");
                }

                Guid pasrseValue;
                if (!Guid.TryParse(idattributeValue, out pasrseValue))
                {
                    invalidIdValues.Add(idattributeValue);
                }

                #endregion Verify ID attribute
            }

            // Add logs for R222
            string logMsg = string.Empty;
            if (invalidIdValues.Count != 0)
            {
                string title = "These below values indicate invalid Id values in [ActiveWorkflowsData] element:";
                logMsg = this.GenerateLogsForMutipleItem(invalidIdValues, title);
                this.Site.Log.Add(LogEntryKind.Debug, logMsg);
            }

            // Verify R223
            this.Site.CaptureRequirementIfAreEqual(
                        0,
                        invalidIdValues.Count,
                        223,
                        @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData. Workflows.Workflow.Id: A workflow identifier.");

            // Add logs for R222
            logMsg = string.Empty;
            if (invalidStatusPageUrlValues.Count != 0)
            {
                string title = "There are invalid StatusPageUrl values:";
                logMsg = this.GenerateLogsForMutipleItem(invalidStatusPageUrlValues, title);
                this.Site.Log.Add(LogEntryKind.Debug, logMsg);
            }

            // Verify R222
            this.Site.CaptureRequirementIfAreEqual(
                         0,
                         invalidStatusPageUrlValues.Count,
                         222,
                         @"[In GetWorkflowDataForItemResponse] This URL[GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData.Workflows.Workflow.StatusPageUrl] MUST be both a fully qualified URL and an absolute URL.");

            // Add the log for R448.
            logMsg = string.Empty;
            if (invalidInternalStateValues.Count != 0)
            {
                string title = "There are invalid InternalState values:";
                logMsg = this.GenerateLogsForMutipleItem(invalidInternalStateValues, title);
                this.Site.Log.Add(LogEntryKind.Debug, logMsg);
            }
 
            if (Common.IsRequirementEnabled(422, this.Site))
            {
                // Add the log for R422.
                logMsg = string.Empty;

                // If the response have any item which does not include ActivityDetails attributes, output the logs.
                if (itemsWithOutActivityDetailsAttribute.Count != 0)
                {
                    string title = "These below ids indicate those workflow items in [ActiveWorkflowsData] element does not include [ActivityDetails] attribute:";
                    logMsg = this.GenerateLogsForMutipleItem(itemsWithOutActivityDetailsAttribute, title);
                    this.Site.Log.Add(LogEntryKind.Debug, logMsg);
                }

                // If all items have ActivityDetails Attribute, then verify R422
                this.Site.CaptureRequirementIfAreEqual(
                           workflows.ChildNodes.Count,
                           itemsWithActivityDetailsAttribute.Count,
                           422,
                           @"[In Appendix B: Product Behavior] Implementation does include his attribute[ActivityDetails].(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(365, this.Site))
            {
                // Add the log for R365.
                logMsg = string.Empty;

                // If the response have any item which include ActivityDetails attributes, output the logs.
                if (itemsWithActivityDetailsAttribute.Count != 0)
                {
                    string title = "These below ids indicate those workflow items in [ActiveWorkflowsData] element does not include [ActivityDetails] attribute:";
                    logMsg = this.GenerateLogsForMutipleItem(invalidInternalStateValues, title);
                    this.Site.Log.Add(LogEntryKind.Debug, logMsg);
                }

                // If all items have ActivityDetails Attribute, then Verify R365
                this.Site.CaptureRequirementIfAreEqual(
                           workflows.ChildNodes.Count,
                           itemsWithOutActivityDetailsAttribute.Count,
                           365,
                           @"[In Appendix B: Product Behavior] Implementation does not include this attribute[ActivityDetails]. [In Appendix B: Product Behavior] <3> Section 3.1.4.5.2.2:  Office SharePoint Server 2007 does not include this attribute[ActivityDetails].");
            }

            if (Common.IsRequirementEnabled(423, this.Site))
            {
                // Add the log for R423.
                logMsg = string.Empty;
                if (itemsWithCorrelationIdAttribute.Count != 0)
                {
                    string title = "These below ids indicate those workflow items in [ActiveWorkflowsData] element does include [CorrelationId] attribute:";
                    logMsg = this.GenerateLogsForMutipleItem(itemsWithCorrelationIdAttribute, title);
                    this.Site.Log.Add(LogEntryKind.Debug, logMsg);
                }

                // Verify R423
                this.Site.CaptureRequirementIfAreEqual(
                           0,
                           itemsWithCorrelationIdAttribute.Count,
                           423,
                           @"[In Appendix B: Product Behavior] Implementation does not include this attribute[CorrelationId]. [In Appendix B: Product Behavior] <4> Section 3.1.4.5.2.2:  Office SharePoint Server 2007 and SharePoint Server 2010 do not include this attribute[CorrelationId].");
            }

            if (Common.IsRequirementEnabled(424, this.Site))
            {
                // Add the log for R424.
                logMsg = string.Empty;
                if (itemsWithOutCorrelationIdAttribute.Count != 0)
                {
                    string title = "These below ids indicate those workflow items in [ActiveWorkflowsData] element does include [CorrelationId] attribute:";
                    logMsg = this.GenerateLogsForMutipleItem(itemsWithOutCorrelationIdAttribute, title);
                    this.Site.Log.Add(LogEntryKind.Debug, logMsg);
                }

                // Verify R424
                this.Site.CaptureRequirementIfAreEqual(
                           0,
                           itemsWithOutCorrelationIdAttribute.Count,
                           424,
                           @"[In Appendix B: Product Behavior] Implementation does include his attribute[CorrelationId].(Microsoft® SharePoint® Foundation 2013 Preview follow this behavior.)");
            }

            #endregion Capture Workflows related requirement
        }

        #endregion CaptureRequirements
    }
}