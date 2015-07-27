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
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Linq;
    using System.Web.Services.Protocols;
    using System.Xml;

    /// <summary>
    /// The TestSuite of MS-WWSP S03 AlterToDo.
    /// </summary>
    [TestClass]
    public class S03_AlterToDo : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region TestCase

        #region MSWWSP_S03_TC01_AlterToDo_Success
        /// <summary>
        /// This test case is used to verify AlterToDo operation, modify the values of Fields on a workflow task successful.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S03_TC01_AlterToDo_Success()
        {
            // Upload a file. 
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start a normal work flow 
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, false);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // initialize the value which will be updated in AlterToDo operation, and this value is unique.
            string alaterValue = this.GenerateRandomValue();

            // initialize a SoapException instance.
            bool isSuccessAlterToDo = false;
            AlterToDoResponseAlterToDoResult alterToDoResult = new AlterToDoResponseAlterToDoResult();
            try
            {
                // Call method AlterToDo to modify the values of fields on a workflow task. 
                XmlElement taskData = this.GetAlerToDoTaskData(this.Site, alaterValue);
                alterToDoResult = ProtocolAdapter.AlterToDo(uploadFileUrl, int.Parse(taskIdValue), new Guid(TaskListId), taskData);
                isSuccessAlterToDo = true;
            }
            catch (SoapException ex)
            {
                isSuccessAlterToDo = false;
                this.Site.Assert.Fail("AlterToDo operation is failed." + ex.ToString());
            }
       
           // Call method GetToDosForItem to get a set of Workflow Tasks for a document. 
            GetToDosForItemResponseGetToDosForItemResult todosAfterStartTask = ProtocolAdapter.GetToDosForItem(uploadFileUrl);
            var matchAttributeitems = from XmlAttribute attributeitem in todosAfterStartTask.ToDoData.xml.data.Any[0].Attributes
                             where attributeitem.Value.Equals(alaterValue, StringComparison.OrdinalIgnoreCase)
                             select attributeitem;

            // If the specified workflow task is be modified, then R70, R77, R83, R415, R93 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R70
            Site.CaptureRequirementIfAreEqual(
                1,
                matchAttributeitems.Count(),
                70,
                @"[In Message Processing Events and Sequencing Rules]The operation AlterToDo modifies a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R77
            Site.CaptureRequirementIfAreEqual(
                1,
                matchAttributeitems.Count(),
                77,
                @"[In AlterToDo] This operation is used to modify the values of Fields on a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R83
            Site.CaptureRequirementIfAreEqual(
                1,
                matchAttributeitems.Count(),
                83,
                @"[In Messages] AlterToDoSoapOut specifies the response to a request to modify the values of Fields on a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R415
            Site.CaptureRequirementIfAreEqual(
                1,
                matchAttributeitems.Count(),
                415,
                @"[In Elements] AlterToDoResponse contains the response to a request to modify the values of Fields on a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R93
            Site.CaptureRequirementIfAreEqual(
                1,
                matchAttributeitems.Count(),
                93,
                @"[In AlterToDo] This element[AlterToDo] is sent with AlterToDoSoapIn and specifies the workflow task to be modified, as well as the fields and values to be modified.");

            // If the response from the AlterToDo operation is not null, then R80 and R102 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R80
            Site.CaptureRequirementIfIsNotNull(
                alterToDoResult,
                80,
                @"[In AlterToDo] The protocol client sends an AlterToDoSoapIn request message, and the protocol server responds with an AlterToDoSoapOut response message.");

            // Verify MS-WWSP requirement: MS-WWSP_R102
            Site.CaptureRequirementIfIsNotNull(
                alterToDoResult,
                102,
                @"[In AlterToDoResponse] This element is sent with AlterToDoSoapOut and specifies whether the AlterToDo operation was successful.");

            // If AlterToDoResult.fSuccess is equal to 1, then R105 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R105
            Site.CaptureRequirementIfAreEqual(
                1,
                alterToDoResult.fSuccess,
                105,
                @"[In AlterToDoResponse] AlterToDoResult.fSuccess: If the operation[AlterToDo] was successful, this[AlterToDoResult.fSuccess] MUST be set to 1.");

            // If there is no soap exception, then R108 and R109 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R108
            bool isVerifyR108 = isSuccessAlterToDo;
            
            Site.CaptureRequirementIfIsTrue(
                isVerifyR108,
                108,
                @"[In AlterToDoResponse]The success of this operation[AlterToDo] MUST NOT substitute for a SOAP fault [or HTTP Status Code] in the case of a protocol server fault.");
        
            // Verify MS-WWSP requirement: MS-WWSP_R109
            bool isVerifyR109 = isSuccessAlterToDo;
            
            Site.CaptureRequirementIfIsTrue(
                isVerifyR109,
                109,
                @"[In AlterToDoResponse] The success of this operation[AlterToDo] MUST NOT substitute for [a SOAP fault or] HTTP Status Code in the case of a protocol server fault.");
        }
        #endregion

        #region MSWWSP_S03_TC02_AlterToDo_Fail
        /// <summary>
        /// This test case is used to verify AlterToDo operation with the invalid document URL, the operation failed.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S03_TC02_AlterToDo_Fail()
        {
            // Call method AlterToDo to modify the values of fields on a workflow task with invalid parameters. 
            AlterToDoResponseAlterToDoResult alterToDoResult = new AlterToDoResponseAlterToDoResult();
            try
            {
                alterToDoResult = ProtocolAdapter.AlterToDo(null, 0, new Guid(), null);
            }
            catch (SoapException ex)
            {
                Site.Log.Add(LogEntryKind.Debug, "AlterToDo is failed when the parameters are invalid, S03_TC02_AlterToDo_Fail" + ex.Detail);
            }

            // If the AlterToDoResult.fSuccess is equal to 0, then R106 should be covered.        
            // Verify MS-WWSP requirement: MS-WWSP_R106
            Site.CaptureRequirementIfAreEqual(
                0,
                alterToDoResult.fSuccess,
                106,
                @"[In AlterToDoResponse] Otherwise[if the AlterToDo operation was failed], this[AlterToDoResult.fSuccess] MUST be set to zero.");
        }

        #endregion

        #region MSWWSP_S03_TC03_AlterToDo_IgnoreItem
        /// <summary>
        /// This test case is used to verify set the different string as the item value, server reply same if the site of the SOAP request URL contains a list with the specified todoListId.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S03_TC03_AlterToDo_IgnoreItem()
        {
            // Upload a file. 
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start a normal work flow 
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, false);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // initialize the value which will be updated in AlterToDo operation, and this value is unique.
            string alaterValueFirst = this.GenerateRandomValue();
          
            // Call method AlterToDo to modify the values of fields on a workflow task. 
            XmlElement taskData = this.GetAlerToDoTaskData(this.Site, alaterValueFirst);
            AlterToDoResponseAlterToDoResult alterToDoResult = ProtocolAdapter.AlterToDo(uploadFileUrl, int.Parse(taskIdValue), new Guid(TaskListId), taskData);

            // Call method GetToDosForItem to get a set of Workflow Tasks for a document. 
            GetToDosForItemResponseGetToDosForItemResult todosAfterStartTask = ProtocolAdapter.GetToDosForItem(uploadFileUrl);
            var matchAttributeitemsFirst = from XmlAttribute attributeitem in todosAfterStartTask.ToDoData.xml.data.Any[0].Attributes
                                      where attributeitem.Value.Equals(alaterValueFirst, StringComparison.OrdinalIgnoreCase)
                                      select attributeitem;

            // initialize an not existing file Url.
            string notExistingFileUrl = this.GenerateRandomValue();

            // initialize the value which will be updated in AlterToDo operation, and this value is unique.
            string alaterValueSecond = this.GenerateRandomValue();

            // Call method AlterToDo to modify the values of fields on a workflow task. 
            taskData = this.GetAlerToDoTaskData(this.Site, alaterValueSecond);
            alterToDoResult = ProtocolAdapter.AlterToDo(notExistingFileUrl, int.Parse(taskIdValue), new Guid(TaskListId), taskData);

            // Call method GetToDosForItem to get a set of Workflow Tasks for a document. 
            todosAfterStartTask = ProtocolAdapter.GetToDosForItem(uploadFileUrl);
            var matchAttributeitemsSecond = from XmlAttribute attributeitem in todosAfterStartTask.ToDoData.xml.data.Any[0].Attributes
                                  where attributeitem.Value.Equals(alaterValueSecond, StringComparison.OrdinalIgnoreCase)
                                      select attributeitem;

            // Verify MS-WWSP requirement: MS-WWSP_R97
            // If the altered value present in response of GetToDosForItem operation, the AlterToDo operation succeed.
            bool isVerifyR97 = matchAttributeitemsFirst.Count() == 1;

            // If the task is modified successfully in the first and second AlterToDo operation, then R97 should be covered.
            isVerifyR97 = isVerifyR97 && matchAttributeitemsSecond.Count() == 1;

            // Verify MS-WWSP requirement: MS-WWSP_R97
            Site.CaptureRequirementIfIsTrue(
                isVerifyR97,
                97,
                @"[In AlterToDo] Set the different string as the item value, server reply same if the site (2) of the SOAP request URL contains a list with the specified todoListId.");
        }
        #endregion

        #endregion
    }
}