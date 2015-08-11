namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-WWSP SUT Controller adapter. It perform the operations to the protocol SUT which are not specified in [MS-WWSP] protocol.
    /// </summary>
    public interface IMS_WWSPSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// A method is used to Upload a file to the specified document library.
        /// </summary>
        /// <param name="documentLibraryTitle">A parameter represents the title of a document library where the file will be uploaded.</param>
        /// <returns>A return value represents the absolute URL of the file on the specified document library if succeed, otherwise return null.</returns>
        [MethodHelp(@"Enter the absolute URL of the uploaded file on the specified document library specified in the ""documentLibraryTitle"" input parameter. Entering null indicates that the upload action has failed.")]
        string UploadFileToDocumentLibrary(string documentLibraryTitle);

        /// <summary>
        /// Get the current web title of the site where the test suite run against
        /// </summary>
        /// <returns>A return value represents the web title get from the protocol SUT.</returns>
        [MethodHelp(@"Enter the title of the website where the test suite run against. Entering null indicates that the get currentWeb title action has failed.")]
        string GetCurrentWebTitle();

        /// <summary>
        /// Get the list id by specified list name
        /// </summary>
        /// <param name="listName">A parameter represents the list name which is used as search condition to get the list id.</param>
        /// <returns>A return value represents the list id of the list which match the specified list name value.</returns>
        [MethodHelp(@"Enter the id (GUID format) of the list whose name is equal to the value specified in the ""listName"" input parameter on the target web site. Entering null indicates that the get listId action has failed.")]
        string GetListIdByName(string listName);

        /// <summary>
        ///  Get the Workflow association Id according to the workflow association name.
        /// </summary>
        /// <param name="targetListName">A parameter represents the list name where the association is located on.</param>
        /// <param name="workFlowAssociationName">A parameter represents the workflow association name which is used as search condition to get the workflow association id.</param>
        /// <returns>A return value represents the workflow association Id which the test suite run against.</returns>
        [MethodHelp(@"Enter the Id (GUID format) of workflowAssociation whose name is equal to the specified value of the ""workFlowAssociationName"" input parameter on the specified DocumentLibrary list. Entering null indicates that the get workflowAssociation ID action has failed.")]
        string GetWorkflowAssociationIdByName(string targetListName, string workFlowAssociationName);

        /// <summary>
        /// Cleans up the uploaded files whose URLs are been specified.
        /// </summary>
        /// <param name="currentDoclibraryName">A parameter represents the list name which is used as search condition to get the list id.</param>
        /// <param name="uploadedfilesUrls">A parameter represents a string which contains all URLs for the uploaded files, separated by ",".</param>
        /// <returns>Returns True indicating Cleanup uploaded files was successful</returns>
        [MethodHelp(@"The uploadedfilesUrls parameter contains all URLs of uploaded files expected to delete and separated by the ','. Enter the clean up result for uploaded files of which URLs are included in the ""uploadedfilesUrls"" string. Entering false indicates that the cleanup uploaded files action has failed.")]
        bool CleanUpUploadedFiles(string currentDoclibraryName, string uploadedfilesUrls);

        /// <summary>
        /// Clean up the workflow tasks started by test suite according to specified task ids.
        /// </summary>
        /// <param name="currentTaskListName">A parameter represents the list name which is used as search condition to get the list id.</param>
        /// <param name="taskIds">A parameter represents a string which contains all task ids, separated by ",".</param>
        /// <returns>Returns True indicating Cleanup all started tasks was successful</returns>
        [MethodHelp(@"The TaskIds parameter contains all ids of tasks expected to delete and separated by ','. Enter the clean up result for started workflow tasks of which taskId are included in the ""taskIds"" string. Entering false indicates that the cleanup started tasks action has failed.")]
        bool CleanUpStartedTasks(string currentTaskListName, string taskIds);

        /// <summary>
        /// Get the list URL by specified list name.
        /// </summary>
        /// <param name="listName">A parameter represents the list name which is used as search condition to get the list URL.</param>
        /// <returns>A return value represents the URL of the list which match the specified list name value.</returns>
        [MethodHelp(@"Enter the URL (absolute or relative format, it depends on the implementation) of the list whose name is equal to the value specified in the ""listName"" input parameter on the target web site. Entering null indicates that the get listId action has failed.")]
        string GetListUrlByName(string listName);

        /// <summary>
        /// Get the current web URL of the web site where the test suite run against
        /// </summary>
        /// <returns>A return value represents the web URL get from the protocol SUT.</returns>
        [MethodHelp(@"Enter the web URL (absolute or relative format, it depends on the implementation) of the web site where the test suite run against. Entering null indicates that the get currentWeb title action has failed.")]
        string GetCurrentWebUrl();

        /// <summary>
        ///  Get the Workflow association base Id according to the workflow association name.
        /// </summary>
        /// <param name="targetListName">A parameter represents the list name where the association is located on.</param>
        /// <param name="workFlowAssociationName">A parameter represents the workflow association name which is used as search condition to get the workflow association id.</param>
        /// <returns>A return value represents the workflow association base Id which the test suite run against.</returns>
        [MethodHelp(@"Enter the base Id (GUID format) of workflowAssociation of which name is equal to the specified value of ""workFlowAssociationName"" input parameter on the specified DocumentLibrary list. Entering null indicates that the get workflowAssociation Id action has failed.")]
        string GetBaseIdOfWorkFlowAssociation(string targetListName, string workFlowAssociationName);
    }
}