namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Contains a folderContent.
    /// SubFolder            = StartSubFld folderContent EndFolder
    /// </summary>
    public class SubFolder : SyntacticalBase
    {
        /// <summary>
        /// The start marker of SubFolder.
        /// </summary>
        public const Markers StartMarker = Markers.PidTagStartSubFld;

        /// <summary>
        /// The end marker of SubFolder.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagEndFolder;

        /// <summary>
        /// A folderContent value contains the content of a folder: its properties, messages, and subfolders.
        /// </summary>
        private FolderContent folderContent;

        /// <summary>
        /// Initializes a new instance of the SubFolder class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public SubFolder(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets the folderContent.
        /// </summary>
        public FolderContent FolderContent
        {
            get
            {
                return this.folderContent;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized SubFolder.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized SubFolder, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(StartMarker);
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (stream.ReadMarker(StartMarker))
            {
                this.folderContent = new FolderContent(stream);
                if (stream.ReadMarker(EndMarker))
                {
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }  
    }
}