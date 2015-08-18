namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Contains a folderContent.
    /// TopFolder            = StartTopFld folderContent EndFolder
    /// </summary>
    public class TopFolder : SyntacticalBase
    {
        /// <summary>
        /// The start marker of TopFolder.
        /// </summary>
        public const Markers StartMarker = Markers.PidTagStartTopFld;

        /// <summary>
        /// The end marker of TopFolder.
        /// </summary>
        public const Markers EndMarker = Markers.PidTagEndFolder;

        /// <summary>
        /// A folderContent value contains the content of a folder: its properties, messages, and subfolders.
        /// </summary>
        private FolderContent folderContent;

        /// <summary>
        /// Initializes a new instance of the TopFolder class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public TopFolder(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the folderContent.
        /// </summary>
        public FolderContent FolderContent
        {
            get { return this.folderContent; }
            set { this.folderContent = value; }
        }

        /// <summary>
        /// Verify a stream's current position contains a serialized TopFolder.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized TopFolder, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagStartTopFld);
        }

        /// <summary>
        /// Get the corresponding AbstractFastTransferStream.
        /// </summary>
        /// <returns>The corresponding AbstractFastTransferStream.</returns>
        public AbstractFastTransferStream GetAbstractFastTransferStream()
        {
            AbstractFastTransferStream afts = default(AbstractFastTransferStream);
            afts.StreamType = FastTransferStreamType.TopFolder;

            // If the PropList is null or not contains any properties' values then the IsFolderPropertiesNotOut is true.
            afts.AbstractTopFolder.AbsFolderContent.IsFolderPropertiesNotOut = !(this.FolderContent.PropList.PropValues != null && this.FolderContent.PropList.PropValues.Count > 0);

            // If the field of WarningCode of FolderContent is not null then  IsNoPermissionObjNotOut is true.
            afts.AbstractTopFolder.AbsFolderContent.IsNoPermissionObjNotOut = this.FolderContent.WarningCode != null;
            afts.AbstractTopFolder.AbsFolderContent.IsPidTagEcWarningOut = this.FolderContent.WarningCode != null;
            afts.AbstractTopFolder.AbsFolderContent.IsPidTagNewFXFolderOut = this.FolderContent.HasNewFXFolder;

            // IsFolderContentNotOut Equivalent to the folderContent contains no message changes.
            // All subfolders do not contains subfolders.
            afts.AbstractTopFolder.AbsFolderContent.IsFolderContentNotOut = !(((this.FolderContent.FolderMessages != null 
                                                        && this.FolderContent.FolderMessages.MessageTupleList != null
                                                        && this.FolderContent.FolderMessages.MessageTupleList.Count > 0
                                                        && this.FolderContent.FolderMessages.MessageTupleList[0].Item2.Messages != null
                                                        && this.FolderContent.FolderMessages.MessageTupleList.Count > 0
                                                        && this.FolderContent.FolderMessages.MessageTupleList[0].Item2.Messages.Count > 0)
                                                        || (this.FolderContent.SubFolders != null && this.FolderContent.SubFolders.Count > 0))
                                                        || (this.folderContent.PropList != null));

            // If contains subFolders.
            if (this.FolderContent.SubFolders != null && this.FolderContent.SubFolders.Count > 0)
            {
                // Contains subFolders.
                afts.AbstractTopFolder.SubFolderInScope = true;

                // Check whether subFolders is Preceded By PidTagFXDelProp.
                if (this.folderContent.PropList.IsNoMetaPropertyContained)
                {
                    for (int i = 0; i < this.folderContent.PropList.PropValues.Count; i++)
                    {
                        PropValue p = this.folderContent.PropList.PropValues[i];
                        if (p.PropInfo.PropID == 0x360E)
                        {
                            afts.AbstractFolderContent.IsSubFolderPrecededByPidTagFXDelProp = true;
                            break;
                        }
                    }
                }
            }

            return afts;
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