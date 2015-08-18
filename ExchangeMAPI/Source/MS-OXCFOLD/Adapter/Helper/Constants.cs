namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    /// <summary>
    /// Constants used in the OXCFOLD.
    /// </summary>
    public class Constants
    {
        /// <summary>
        /// The protocol name of "MS-OXCDATA"
        /// </summary>
        public const string MSOXCDATA = "MS-OXCDATA";

        /// <summary>
        /// The protocol name of "MS-OXPROPS"
        /// </summary>
        public const string MSOXPROPS = "MS-OXPROPS";

        /// <summary>
        /// The common logon ID used in connection.
        /// </summary>
        public const byte CommonLogonId = 0x00;

        /// <summary>
        /// The common InputHandleIndex used in request message.
        /// </summary>
        public const byte CommonInputHandleIndex = 0x00;

        /// <summary>
        /// The common OutputHandleIndex used in request message.
        /// </summary>
        public const byte CommonOutputHandleIndex = 0x01;

        /// <summary>
        /// The successful code which the server returned.
        /// </summary>
        public const int SuccessCode = 0x00000000;

        /// <summary>
        /// The null character terminate a 'Null-terminated' string.
        /// </summary>
        public const string StringNullTerminated = "\0";

        /// <summary>
        /// The index of public folder of Exchange Server. Specified in section 2.2.1.1.3 in [MS-OXCSTOR].
        /// </summary>
        public const int PublicFolderIndex = 1;

        /// <summary>
        /// The index of inbox of Exchange Server. Specified in section 2.2.1.1.3 in [MS-OXCSTOR].
        /// </summary>
        public const int InboxIndex = 4;

        #region Folder and Message Name.

        /// <summary>
        /// The name prefix of the root folder which is used by MS-OXCFOLD test cases.
        /// This folder created for MS-OXCFOLD test suite under which all other folders
        /// and messages for test will be created.
        /// </summary>
        public const string RootFolder = "MSOXCFOLDRootFolder";

        /// <summary>
        /// SearchFolder1's name which is created by test case as a search folder under the root Folder.
        /// </summary>
        public const string SearchFolder = "MSOXCFOLDSearchFolder1" + Constants.StringNullTerminated;

        /// <summary>
        /// SearchFolder2's name which is created by test case as a search folder under the root Folder.
        /// </summary>
        public const string SearchFolder2 = "MSOXCFOLDSearchFolder2" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder1's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder1 = "MSOXCFOLDSubfolder1" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder2's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder2 = "MSOXCFOLDSubfolder2" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder3's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder3 = "MSOXCFOLDSubfolder3" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder4's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder4 = "MSOXCFOLDSubfolder4" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder5's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder5 = "MSOXCFOLDSubfolder5" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder6's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder6 = "MSOXCFOLDSubfolder6" + Constants.StringNullTerminated;

        /// <summary>
        /// Subfolder7's name which is created by test case as a generic folder under the root Folder.
        /// </summary>
        public const string Subfolder7 = "MSOXCFOLDSubfolder7" + Constants.StringNullTerminated;

        #endregion
    }
}