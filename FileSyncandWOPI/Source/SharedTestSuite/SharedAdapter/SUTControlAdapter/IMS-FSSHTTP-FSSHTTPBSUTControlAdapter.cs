//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// SUT control adapter interface.
    /// </summary>
    public interface IMS_FSSHTTP_FSSHTTPBSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to set the maximum number of clients allowed to join a coauthoring session.
        /// </summary>
        /// <param name="count">The maximum number of clients allowed to join a co-authoring session.</param>
        /// <returns>Return true if set max number of clients succeed, otherwise return false.</returns>
        [MethodHelp(@"Change the maximum number of clients allowed to join a coauthoring session to the specified parameter(count)." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool SetMaxNumOfCoauthUsers(int count);

        /// <summary>
        /// This method is used to check in the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked in.</param>
        /// <param name="userName">Specify the name of the user who checks in the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <param name="checkInComments">Specify the checked in comments.</param>
        /// <returns>Return true if the check in succeeds, otherwise return false.</returns>
        [MethodHelp(@"Check in the file (fileUrl) using the credential (userName, password and domain) and comments (checkInComments)." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool CheckInFile(string fileUrl, string userName, string password, string domain, string checkInComments);

        /// <summary>
        /// This method is used to check out the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked out.</param>
        /// <param name="userName">Specify the name of the user who checks out the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <returns>Return true if the check out succeeds, otherwise return false.</returns>
        [MethodHelp(@"Check out the file (fileUrl) using the credential (userName, password and domain)." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool CheckOutFile(string fileUrl, string userName, string password, string domain);

        /// <summary>
        /// This method is used to change the status of the document library whether needs to check out the files before editing or locking.
        /// </summary>
        /// <param name="isCheckoutRequired">Indicating whether the doc library requires checking out files or not.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"If the specified parameter (isCheckoutRequired) is true, change the status of the document library to require file checkout before editing or locking; otherwise, change the status of the document library to not require file checkout." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool ChangeDocLibraryStatus(bool isCheckoutRequired);

        /// <summary>
        /// This method is used to change the status of the document library whether the coauthoring feature is disabled. 
        /// </summary>
        /// <param name="isDisabled">Specify whether disable the coauthoring feature. If true then disable the coauthoring feature, otherwise enable the coauthoring feature.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"If the specified parameter (isDisabled) is true, disable the coauthoring feature; otherwise, enable the feature." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool SwitchCoauthoringFeature(bool isDisabled);

        /// <summary>
        /// This method is used to get the GUID of the specified list name.
        /// </summary>
        /// <param name="listName">A specified list name in the server.</param>
        /// <returns>The GUID of the list.</returns>
        [MethodHelp("Enter the GUID for the specified listName parameter.")]
        string GetListGuidByName(string listName);

        /// <summary>
        /// This method is used to turn on/turn off the cell storage service.
        /// </summary>
        /// <param name="isEnabled">Specify whether the cell storage service is turned on or turned off. True indicates the cell storage service is turned on, otherwise the service is turned off.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>        
        [MethodHelp(@"If the specified parameter (isEnabled) is true, turn on the cell storage service; otherwise, turn off the service." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool SwitchCellStorageService(bool isEnabled);

        /// <summary>
        /// This method is used to change the authentication mode to windows/claims based authentication.
        /// </summary>
        /// <param name="isClaimsAuthentication">Specify the authentication mode. True indicates the claims based authentication, false indicates the windows based authentication.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"If the specified parameter (isClaimsAuthentication) is true, change the authentication mode to claims-based authentication; otherwise, change the authentication mode to Windows authentication." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool SwitchClaimsAuthentication(bool isClaimsAuthentication);

        /// <summary>
        /// This method is used to turn on and turn off the versioning of the document library.
        /// </summary>
        /// <param name="documentLibraryName">Specify the document library name.</param>
        /// <param name="isEnable">Whether turn on or turn off the versioning of the document library. True indicates turning on the versioning of this document library, false indicates turning off the versioning.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        [MethodHelp(@"If the specified parameter (isEnabled) is true, turn on the versioning feature of the specified (documentLibraryName); otherwise, turn off versioning." +
                      "If the operation succeeds, enter true; otherwise, enter false.")]
        bool SwitchMajorVersioning(string documentLibraryName, bool isEnable);
    }
}