namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter interface.
    /// </summary>
    public interface IMS_LISTSWSSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// A method used to upload a file whose content is generated in the format "MSLISTSWSTEST Test on [HHmmss_fff]" to the specified document library.
        /// </summary>
        /// <param name="documentLibraryTitle">A parameter represents the title of list where the file will be uploaded.</param>
        /// <returns>A return value represents the absolute URL of the file on the specified document library if succeed, otherwise return null.</returns>       
        [MethodHelp(@"Create a text file with a unique file name. The file can contain any content. Upload the file to the specified document library (documentLibraryTitle). Enter the absolute URL of the file on the specified document library (documentLibraryTitle). Enter null if the operation cannot be completed, for e.g. the specified list doesn't exist.")]
        string UploadFile(string documentLibraryTitle);

        /// <summary>
        /// A method used to move a file from source location to a destination location in the same site. File at the destination location will be overwritten if it already exists.
        /// </summary>
        /// <param name="sourceUrl">URL of the file source location. Can be either absolute URL or site relative URL.</param>
        /// <param name="destinationUrl">URL of the location where the file is moved. Can be either absolute URL or site relative URL.</param>
        [MethodHelp(@"Move the specified file(sourceUrl) to the specified location(destinationUrl). The file at the destination location should be overwritten if it already exists.")]
        void MoveFile(string sourceUrl, string destinationUrl);

        /// <summary>
        /// A method used to set Custom Send To Destination Name and Url for document library.
        /// </summary>
        /// <param name="documentLibraryId">The Guid of the document library.</param>
        /// <param name="sendToDestinationName">Custom Send To Destination Name to set.</param>
        /// <param name="sendToDestinationUrl">Custom Send To Destination Url to set</param>
        [MethodHelp(@"Set the Custom Send To Destination Name(sendToDestinationName) and Custom Send To Destination Url(sendToDestinationUrl) for a document library (documentLibraryId).")]
        void SetSendToNameAndUrl(string documentLibraryId, string sendToDestinationName, string sendToDestinationUrl);

        /// <summary>
        /// A method used to get the value of Presence setting of WebApp.
        /// </summary>
        /// <returns>A return value represents the value of Presence Settings of WebApp.</returns>
        [MethodHelp(@"Get the value of the presence setting (On or Off) of the web application and enter the appropriate value in action results. Accepted values are either True or False. Enter ""True"" if the presence setting is on and ""False"" if the presence setting is off.")]
        bool GetWebAppPresence();

        /// <summary>
        /// A method used to get the value of Recycle Bin setting of WebApp.
        /// </summary>
        /// <returns>A return value represents the value of Recycle Bin setting of WebApp.</returns>
        [MethodHelp(@"Get the Recycle Bin setting (On or Off) of the web application and enter the appropriate value in action results. Accepted values are True or false, enter ""True"" if the setting is on & ""False"" if the setting is off.")]
        bool GetWebAppRecycleBin();

        /// <summary>
        /// A method used to set the value of Presence setting of WebApp.
        /// </summary>
        /// <param name="isEnabled">A parameter represents if the Presence is enabled.</param>
        [MethodHelp(@"Set the value of the presence setting(isEnabled) of web application.")]
        void SetWebAppPresence(bool isEnabled);

        /// <summary>
        /// A method used to get the value of Recycle Bin setting of WebApp.
        /// </summary>
        /// <param name="isEnabled">A parameter represents if the Recycle Bin is enabled.</param>
        [MethodHelp(@"Set the value of the Recycle Bin setting(isEnabled) of the web application.")]
        void SetWebAppRecycleBin(bool isEnabled);

        /// <summary>
        /// A method used to Set MajorVersionLimit and MajorWithMinorVersionsLimit value when versioning is enabled in the list.
        /// </summary>
        /// <param name="listId">The id of the list used for this protocol testing.</param>
        /// <param name="majorVersionLimitValue">The int value of MajorVersionLimit.</param>
        /// <param name="majorWithMinorVersionsLimitValue">The int value of MajorWithMinorVersionsLimit.</param>
        /// <returns>A Boolean indicates whether the operation is executed successfully, 
        /// TRUE means the operation is executed successfully, FALSE means not.</returns>
        [MethodHelp("Set MajorVersionLimit (majorVersionLimitValue) and MajorWithMinorVersionsLimit (majorWithMinorVersionsLimitValue) " +
            "value when versioning is enabled in the specified list (listId). The return value should be a Boolean value that indicates whether" +
            " the operation was run successfully. TRUE means the operation was run successfully, " +
            "FALSE means the operation failed.")]
        bool SetVersionLimit(string listId, int majorVersionLimitValue, int majorWithMinorVersionsLimitValue);

        /// <summary>
        /// A method used to Get RootFolder value from the list.
        /// </summary>
        /// <param name="listName">The name of the list used for this protocol testing.</param>
        /// <returns>A string indicates the RootFolder value in the specified list.</returns>
        [MethodHelp("Get RootFolder value from the specified list (listName). The return value should be " +
            "a string value that indicates the RootFolder value in the specified list.")]
        string GetListRootFolder(string listName);
    }
}