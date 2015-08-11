namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-WOPI managed code SUT control adapter, it supports a feature to get the related WOPI resource URL.
    /// </summary>
    public interface IMS_WOPIManagedCodeSUTControlAdapter : IAdapter
    {           
        /// <summary>
        /// A method used to get the WOPI resource URL by specified user credentials.
        /// </summary>
        /// <param name="absoluteUrlOfResource">A parameter represents the absolute URL of normal resource which will be used to get a WOPI resource URL.</param>
        /// <param name="rootResourceType">A parameter indicating the WOPI root resource URL type will be returned.</param>
        /// <param name="userName">A parameter represents the name of user whose associated token will be returned in the WOPI resource URL.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>A return value represents the WOPI resource URL, which can be used in MS-WOPI operations.</returns>
        [MethodHelp(@"Enter the URL of the specified resource (absoluteUrlOfResource) for the specified user account(userName, password, domain). If the 'rootResourceType' parameter presents the value 'FolderLevel' or '0', the expected URL is for folder-level format[HTTP://server/<...>/wopi*/folders/<id>]. If the 'rootResourceType' parameter presents the value 'FileLevel' or '1', the expected URL is for file-level format[HTTP://server/<...>/wopi*/files/<id>]")]
        string GetWOPIRootResourceUrl(string absoluteUrlOfResource, WOPIRootResourceUrlType rootResourceType, string userName, string password, string domain);
    }
}