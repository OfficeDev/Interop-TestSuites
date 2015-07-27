//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS-OXWSSYNC SUT control adapter interface.
    /// </summary>
    public interface IMS_OXWSSYNCSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified meeting message in the Inbox folder, then accept it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="itemSubject">Subject of the meeting message which should be accepted.</param>
        /// <param name="itemType">Type of the meeting message which should be accepted.</param>
        /// <returns>If the specified meeting message is accepted successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and find the (itemType) type meeting message with the subject (itemSubject) in the Inbox folder," +
            " then accept the meeting message.\n" +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool FindAndAcceptMeetingMessage(string userName, string userPassword, string userDomain, string itemSubject, string itemType);

        /// <summary>
        /// Log on to a mailbox with a specified user account and check whether the specified item exists.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified item.</param>
        /// <param name="itemSubject">Subject of the item which should exist.</param>
        /// <param name="itemType">Type of the item which should exist.</param>
        /// <returns>If the item exists, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and check whether the (itemType) type item with the subject (itemSubject)" +
            " exists in the (folderName) folder." +
            " If yes, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool IsItemExisting(string userName, string userPassword, string userDomain, string folderName, string itemSubject, string itemType);

        /// <summary>
        /// Log on to a mailbox with a specified user account and check whether the specified calendar item is cancelled or not. 
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified item.</param>
        /// <param name="itemSubject">Subject of the item which should be canceled.</param>
        /// <returns>If the specified item is canceled, return true, otherwise return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and check whether the item with the subject (itemSubject)" +
            " is cancelled in the (folderName) folder." +
            " If yes, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool IsCalendarItemCanceled(string userName, string userPassword, string userDomain, string folderName, string itemSubject);

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified item then delete it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="folderName">Name of the folder which should be searched for the specified item.</param>
        /// <param name="itemSubject">Subject of the item which should be deleted.</param>
        /// <param name="itemType">Type of the item which should be deleted.</param>
        /// <returns>If the specified item is deleted successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and find the (itemType) type item with the subject (itemSubject) in the (folderName) folder," +
            " then delete it.\n" +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool FindAndDeleteItem(string userName, string userPassword, string userDomain, string folderName, string itemSubject, string itemType);

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified folder then update the folder name of it.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="currentFolderName">Current name of the folder which should be updated.</param>
        /// <param name="newFolderName">New name of the folder which should be updated to.</param>
        /// <returns>If the name of the folder is updated successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and find the (currentFolderName) folder under the (parentFolderName) folder," +
            " then rename the (currentFolderName) folder to (newFolderName).\n" +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool FindAndUpdateFolderName(string userName, string userPassword, string userDomain, string parentFolderName, string currentFolderName, string newFolderName);

        /// <summary>
        /// Log on to a mailbox with a specified user account and find the specified folder, then delete it if it is found.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <param name="parentFolderName">Name of the parent folder.</param>
        /// <param name="subFolderName">Name of the folder which should be updated.</param>
        /// <returns>If the folder is deleted successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName, userPassword, userDomain) and find the (subFolderName) folder under the (parentFolderName) folder,  and if the folder is found, delete it." +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool FindAndDeleteSubFolder(string userName, string userPassword, string userDomain, string parentFolderName, string subFolderName);

        /// <summary>
        /// Log on to a mailbox with a specified user account and delete all the items and subfolders from Inbox, Sent Items, Calendar, Contacts, Tasks, Deleted Items and Search Folders if any.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="userPassword">Password of the user.</param>
        /// <param name="userDomain">Domain of the user.</param>
        /// <returns>If the mailbox is cleaned up successfully, return true; otherwise, return false.</returns>
        [MethodHelp("Log on to a mailbox with a specified user account(userName,userPassword,userDomain) and delete all the items and subfolders from Inbox, Sent Items, Calendar, Contacts, Tasks, Deleted Items, and Search Folders if any.\n" +
            " If the operation succeeds, enter \"TRUE\";" +
            " otherwise, enter \"FALSE\".")]
        bool CleanupMailBox(string userName, string userPassword, string userDomain);
    }
}