//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUTControlAdapter's definition.
    /// </summary>
    public interface IMS_OUTSPSSUTControlAdapter : IAdapter
    {
        #region Interact with ListsService

        /// <summary>
        /// A method used to add a file to the specified document library.
        /// </summary>
        /// <param name="documentLibraryTitle">A parameter represents the title of a document library where the file will be uploaded.</param>
        /// <param name="fileName">A parameter represents the name of the file uploaded to SUT.</param>
        /// <returns>A return value represents the absolute URL of the file on the specified document library if succeed, otherwise return null.</returns>
        [MethodHelp(@"Enter the absolute URL of the file with the specified file name on the specified document library specified in the ""documentLibraryTitle"" input parameter. The file name on the SUT must match the specified file name. Entering null indicates that the upload action has failed.")]
        string AddOneFileToDocumentLibrary(string documentLibraryTitle, string fileName);

        /// <summary>
        /// A method used to upload a file into the specified folder of a document library on the SUT.
        /// </summary>
        /// <param name="documentLibraryTitle">A parameter represents the title of a document library where the file will be uploaded.</param>
        /// <param name="subfolderName">A parameter represent the name of the folder where the file will be uploaded.</param>
        /// <param name="fileName">A parameter represents the file name which is used for new upload file.</param>
        /// <returns>A return value represents the absolute URL of the file which is uploaded to the specified folder of a document library on the SUT if succeed, otherwise return null.</returns>
        [MethodHelp(@"Enter the absolute URL of the uploaded file on the specified folder of the document library specified in the ""documentLibraryTitle"" input parameter. The file name on the SUT must match the specified file name. Entering null indicates that the upload action has failed.")]
        string UploadFileWithFolder(string documentLibraryTitle, string subfolderName, string fileName);

        /// <summary>
        /// A method used to delete the specified folder by the folder name.
        /// </summary>
        /// <param name="listTitle">A parameter represents the title of a list.</param>
        /// <param name="subfolderName">A parameter represents the name of the folder excepted to delete.</param>
        /// <returns>Returns 'true' indicating delete folder was successful.</returns> 
        [MethodHelp(@"Enter 'true' if the delete folder operation is successful. Otherwise enter 'false'.")]
        bool DeleteFolder(string listTitle, string subfolderName);

        #endregion
    }
}