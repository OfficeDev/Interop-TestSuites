//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter's implementation.
    /// </summary>
    public partial class MS_WWSPAdapter : ManagedAdapterBase, IMS_WWSPAdapter
    {
        #region Variables

        /// <summary>
        /// Web service proxy generated from the full WSDL of MS-WWSP protocol
        /// </summary>
        private WorkflowSoap wwspProxy;

        #endregion Variables

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize method, to set default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">A parameter represents a ITestSite instance which is used to get/operate current test suite context.</param>
        public override void Initialize(ITestSite testSite)
        {
            // Set the protocol name of current test suite
            testSite.DefaultProtocolDocShortName = "MS-WWSP";

            base.Initialize(testSite);
            AdapterHelper.Initialize(testSite);

            // Merge the common configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);
            Common.CheckCommonProperties(this.Site, true);

            // Load SHOULDMAY configuration
            Common.MergeSHOULDMAYConfig(this.Site);
           
            // Initialize the proxy.
            this.wwspProxy = Proxy.CreateProxy<WorkflowSoap>(this.Site);

            // Set soap version according to the configuration file.
            this.SetSoapVersion(this.wwspProxy);

            // Setup the request URL.
            this.wwspProxy.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);

            // Configure the service timeout.
            int soapTimeOut = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in minute.
            this.wwspProxy.Timeout = soapTimeOut * 60000;

            // Set Credentials information
            string userName = Common.GetConfigurationPropertyValue("MSWWSPTestAccount", this.Site);
            string password = Common.GetConfigurationPropertyValue("MSWWSPTestAccountPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.wwspProxy.Credentials = new NetworkCredential(userName, password, domain);
           
            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
                Common.AcceptServerCertificate();
            }
        }

        #endregion

        #region Implement IWWSPAdapter

        /// <summary>
        /// This operation is used to get a set of workflow associations for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents a set of workflow associations for specified document item </returns>
        public GetTemplatesForItemResponseGetTemplatesForItemResult GetTemplatesForItem(string item)
        {
            GetTemplatesForItemResponseGetTemplatesForItemResult getTemplatesForItemResult = null;
          
            try
            {
                getTemplatesForItemResult = this.wwspProxy.GetTemplatesForItem(item);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
           
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaGetTemplatesForItem();
            this.CaptureSoapInfoGetTemplatesForItem();

            if (getTemplatesForItemResult != null)
            {
                this.CaptureTemplateDataRelatedRequirements(getTemplatesForItemResult.TemplateData);
            }

            return getTemplatesForItemResult;
        }

        /// <summary>
        /// This operation is used to get a set of workflow tasks for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents a set of  workflow tasks for specified document item </returns>
        public GetToDosForItemResponseGetToDosForItemResult GetToDosForItem(string item)
        {
            GetToDosForItemResponseGetToDosForItemResult getToDosForItemResult = null;
            
            try
            {
                getToDosForItemResult = this.wwspProxy.GetToDosForItem(item);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
         
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaGetToDosForItem();
            this.CaptureSoapInfoGetToDosForItem();

            if (getToDosForItemResult != null)
            {
                this.CaptureToDoDataRelatedRequirements(getToDosForItemResult.ToDoData);
            }

            return getToDosForItemResult;
        }

        /// <summary>
        /// This operation is used to query a set of workflow associations, workflow tasks, and workflows for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents the WorkflowDatas include workflow associations, workflow tasks, and workflows.</returns>
        public GetWorkflowDataForItemResponseGetWorkflowDataForItemResult GetWorkflowDataForItem(string item)
        {
            GetWorkflowDataForItemResponseGetWorkflowDataForItemResult getWorkflowDataForItemResult = null;
             
            try
            {
                getWorkflowDataForItemResult = this.wwspProxy.GetWorkflowDataForItem(item);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
          
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaGetWorkflowDataForItem();
            if (getWorkflowDataForItemResult != null)
            {   
                this.CaptureGetWorkflowDataForItem(getWorkflowDataForItemResult);
                this.CaptureToDoDataRelatedRequirements(getWorkflowDataForItemResult.WorkflowData.ToDoData);
            }
          
            return getWorkflowDataForItemResult;
        }

        /// <summary>
        /// This operation is used to start a new workflow task, it generating a workflow task base on specified workflow association.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="templateId">A parameter represents a GUID of a workflow association which the workflow task will base on.</param>
        /// <param name="workflowParameters">A parameter represents XML contents to be used by the workflow upon creation. And the contents of this element is considered vendor-extensible</param>
        /// <returns>A return value represents the response data of StartWorkflow operation. This element is unused and the protocol client MUST ignore this element.</returns>
        public object StartWorkflow(string item, Guid templateId, XmlNode workflowParameters)
        {
            object startResult = null;
             
            try
            {
                startResult = this.wwspProxy.StartWorkflow(item, templateId, workflowParameters);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
        
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaStartWorkflow();
            this.CaptureSoapInfoStartWorkflow();
            return startResult;
        }

        /// <summary>
        /// This operation is used to retrieve data about a single workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item which is used as condition to search where the task start.</param>
        /// <param name="taskId">A parameter represents an integer which is the id of task item in a task type list, which is specified by workflow association setting.</param>
        /// <param name="listId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <returns>A return value represents the WorkflowTaskData for the specified single task item.</returns>
        public GetWorkflowTaskDataResponseGetWorkflowTaskDataResult GetWorkflowTaskData(string item, int taskId, Guid listId)
        {
            GetWorkflowTaskDataResponseGetWorkflowTaskDataResult getWorkflowTaskDataResult = null;
             
            try
            {
                getWorkflowTaskDataResult = this.wwspProxy.GetWorkflowTaskData(item, taskId, listId);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
       
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaGetWorkflowTaskData();
            this.CaptureSoapInfoGetWorkflowTaskData();
            return getWorkflowTaskDataResult;
        }

        /// <summary>
        /// This operation is used to modify the values of fields for a workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="todoId">A parameter represents the Id of a task item which is identifying a workflow task to be modified.</param>
        /// <param name="todoListId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <param name="taskData">A parameter represents a set of elements representing field names and values to be altered on a workflow task.</param>
        /// <returns>A return value represents the alterTodo operation execution result</returns>
        public AlterToDoResponseAlterToDoResult AlterToDo(string item, int todoId, Guid todoListId, XmlElement taskData)
        {   
            // Sleep a specified time span in order to the started tasks is in modifiable.
            int delayValue = Common.GetConfigurationPropertyValue<int>("DelayTimeBeforeModifyTask", this.Site);
            if (delayValue > 0)
            {
                delayValue = delayValue * 1000;
                System.Threading.Thread.Sleep(delayValue);
            }

            AlterToDoResponseAlterToDoResult alterToDoResult = null;
             
            try
            {
                alterToDoResult = this.wwspProxy.AlterToDo(item, todoId, todoListId, taskData);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
       
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaAlterToDo();
            this.CaptureSoapInfoAlterToDo();
            return alterToDoResult;
        }

        /// <summary>
        /// This operation is used to claim or release a workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="taskId">A parameter represents the Id of a task item which is identifying a workflow task to be claim or release.</param>
        /// <param name="listId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <param name="isClaim">A parameter represents a bool value to indicate whether the operation is a claim or a release.</param>
        /// <returns>A return value represents the execution result of ClaimReleaseTask operation, include some data info for the operation execution.</returns>
        public ClaimReleaseTaskResponseClaimReleaseTaskResult ClaimReleaseTask(string item, int taskId, Guid listId, bool isClaim)
        {
            // Add a delay time before change the task instance.
            int delayValue = Common.GetConfigurationPropertyValue<int>("DelayTimeBeforeModifyTask", this.Site);
            if (delayValue > 0)
            {
                delayValue = delayValue * 1000;
                System.Threading.Thread.Sleep(delayValue);
            }

            ClaimReleaseTaskResponseClaimReleaseTaskResult claimReleaseTaskResult = null;
             
            try
            {
                claimReleaseTaskResult = this.wwspProxy.ClaimReleaseTask(item, taskId, listId, isClaim);
            }
            catch (SoapException)
            {
                this.CaptureHTTPStatusAndSoapFaultRequirements();
                throw;
            }
         
            // Capture requirements
            this.CaptureTranstportAndSOAPRequirements();
            this.CaptureXMLSchemaClaimReleaseTask();
            this.CaptureSoapInfoClaimReleaseTask();

            return claimReleaseTaskResult;
        }

        #endregion Implement IWWSPAdapter

        #region Private helper methods

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        /// <param name="proxyInstance">A parameter represents the proxy instance which will be set soap version</param>
        private void SetSoapVersion(WorkflowSoap proxyInstance)
        {
            string soapVersion = Common.GetConfigurationPropertyValue("SoapVersion", this.Site);

            if (string.Compare(soapVersion, "SOAP11", true) == 0)
            {
                proxyInstance.SoapVersion = SoapProtocolVersion.Soap11;
            }
            else if (string.Compare(soapVersion, "SOAP12", true) == 0)
            {
                proxyInstance.SoapVersion = SoapProtocolVersion.Soap12;
            }
            else
            {
                Site.Assume.Fail(
                    "Property SoapVersion value must be {0} or {1} at the ptfconfig file.",
                    "SOAP11",
                    "SOAP12");
            }
        }

        /// <summary>
        /// Generate a log message by specified title and list all detail items.
        /// </summary>
        /// <param name="detialItems">A parameter represents all  detail items.</param>
        /// <param name="title">A parameter represents the title of the log message.</param>
        /// <returns>A return value represents the log message.</returns>
        private string GenerateLogsForMutipleItem(List<string> detialItems, string title)
        {
            string logMessage = string.Empty;
            StringBuilder strBuilder = new StringBuilder();
            strBuilder.AppendLine(title);
            if (detialItems != null)
            {
                foreach (string detial in detialItems)
                {
                    strBuilder.AppendLine(detial);
                }
            }

            logMessage = strBuilder.ToString();
            return logMessage;
        }

        #endregion
    }
}