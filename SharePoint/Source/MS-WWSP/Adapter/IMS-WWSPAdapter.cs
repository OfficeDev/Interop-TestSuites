namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Xml;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-WWSP adapter.
    /// </summary>
    public interface IMS_WWSPAdapter : IAdapter
    {
        /// <summary>
        /// This operation is used to get a set of workflow associations for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents a set of workflow associations for specified document item </returns>
        GetTemplatesForItemResponseGetTemplatesForItemResult GetTemplatesForItem(string item);

        /// <summary>
        /// This operation is used to get a set of workflow tasks for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents a set of  workflow tasks for specified document item </returns>
        GetToDosForItemResponseGetToDosForItemResult GetToDosForItem(string item);

        /// <summary>
        /// This operation is used to query a set of workflow associations, workflow tasks, and workflows for specified document item in a document library.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <returns>A return value represents the WorkflowDatas include workflow associations, workflow tasks, and workflows.</returns>
        GetWorkflowDataForItemResponseGetWorkflowDataForItemResult GetWorkflowDataForItem(string item);

        /// <summary>
        /// This operation is used to start a new workflow task, it generating a workflow task base on specified workflow association.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="templateId">A parameter represents a GUID of a workflow association which the workflow task will base on.</param>
        /// <param name="workflowParameters">A parameter represents XML contents to be used by the workflow upon creation. And the contents of this element is considered vendor-extensible</param>
        /// <returns>A return value represents the response data of StartWorkflow operation. This element is unused and the protocol client MUST ignore this element.</returns>
        object StartWorkflow(string item, Guid templateId, XmlNode workflowParameters);

        /// <summary>
        /// This operation is used to retrieve data about a single workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item which is used as condition to search where the task start.</param>
        /// <param name="taskId">A parameter represents an integer which is the id of task item in a task type list, which is specified by workflow association setting.</param>
        /// <param name="listId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <returns>A return value represents the WorkflowTaskData for the specified single task item.</returns>
        GetWorkflowTaskDataResponseGetWorkflowTaskDataResult GetWorkflowTaskData(string item, int taskId, Guid listId);

        /// <summary>
        /// This operation is used to modify the values of fields for a workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="todoId">A parameter represents the Id of a task item which is identifying a workflow task to be modified.</param>
        /// <param name="todoListId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <param name="taskData">A parameter represents a set of elements representing field names and values to be altered on a workflow task.</param>
        /// <returns>A return value represents the alterTodo operation execution result</returns>
        AlterToDoResponseAlterToDoResult AlterToDo(string item, int todoId, Guid todoListId, XmlElement taskData);

        /// <summary>
        /// This operation is used to claim or release a workflow task.
        /// </summary>
        /// <param name="item">A parameter represents a URL which is point to a document item in a document library.</param>
        /// <param name="taskId">A parameter represents the Id of a task item which is identifying a workflow task to be claim or release.</param>
        /// <param name="listId">A parameter represents the list id (GUID format) of the workflow task list which include the specified task list.</param>
        /// <param name="isClaim">A parameter represents a bool value to indicate whether the operation is a claim or a release.</param>
        /// <returns>A return value represents the execution result of ClaimReleaseTask operation, include some data info for the operation execution.</returns>
        ClaimReleaseTaskResponseClaimReleaseTaskResult ClaimReleaseTask(string item, int taskId, Guid listId, bool isClaim);
    }
}