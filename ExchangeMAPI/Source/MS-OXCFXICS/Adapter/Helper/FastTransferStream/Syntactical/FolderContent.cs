namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// The folderContent element contains the content of a folder: 
    /// its properties, messages, and subFolders.
    ///  folderContent        = propList [PidTagEcWarning]
    ///         ( PidTagNewFXFolder / folderMessages ) 
    ///        [ PidTagFXDelProp *SubFolder ]
    /// Actual stream deserialization:
    /// folderContent        =  propList [PidTagEcWarning]
    ///        ( (*(*PidTagFXDelProp PidTagNewFXFolder)) / folderMessages ) 
    ///       [ *PidTagFXDelProp *SubFolder ]
    /// </summary>
    public class FolderContent : SyntacticalBase
    {
        #region Members
        /// <summary>
        /// Contains the properties of the Folder object, which are possibly affected by property filters.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// The folderMessages element contains the messages contained in a folder.
        /// </summary>
        private FolderMessages folderMessages;

        /// <summary>
        /// The folderContent element contains the content of a folder: its properties, messages, and subFolders.
        /// </summary>
        private List<SubFolder> subFolders;

        /// <summary>
        /// The warning code.
        /// </summary>
        private uint? warningCode;

        /// <summary>
        /// The new fxFolder list.
        /// </summary>
        private List<Tuple<List<uint>, FolderReplicaInfo>> newFXFolderList;

        /// <summary>
        /// The fxdel prop list.
        /// </summary>
        private List<uint> fxdelPropList;

        #region Constructors
        /// <summary>
        /// Initializes a new instance of the FolderContent class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public FolderContent(FastTransferStream stream)
            : base(stream)
        {
        }
        #endregion

        /// <summary>
        /// Gets or sets the NewFXFolderList.
        /// </summary>
        public List<Tuple<List<uint>, FolderReplicaInfo>> NewFXFolderList
        {
            get { return this.newFXFolderList; }
            set { this.newFXFolderList = value; }
        }

        /// <summary>
        /// Gets or sets warningCode.
        /// </summary>
        public uint? WarningCode
        {
            get { return this.warningCode; }
            set { this.warningCode = value; }
        }

        /// <summary>
        /// Gets or sets FXDelPropList before subFolder list.
        /// </summary>
        public List<uint> FXDelPropList
        {
            get { return this.fxdelPropList; }
            set { this.fxdelPropList = value; }
        }

        /// <summary>
        /// Gets or sets folderMessages.
        /// </summary>
        public FolderMessages FolderMessages
        {
            get { return this.folderMessages; }
            set { this.folderMessages = value; }
        }

        /// <summary>
        /// Gets or sets subFolders.
        /// </summary>
        public List<SubFolder> SubFolders
        {
            get { return this.subFolders; }
            set { this.subFolders = value; }
        }

        /// <summary>
        /// Gets or sets propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets a value indicating whether contains pidtagNewfxfolder.
        /// </summary>
        public bool HasNewFXFolder
        {
            get
            {
                return (this.NewFXFolderList != null)
                    && this.NewFXFolderList.Count > 0;
            }
        }
        #endregion
        #region Static methods
        /// <summary>
        /// Verify that a stream's current position contains a serialized folderContent.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized folderContent, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return PropList.Verify(stream);
        }
        #endregion

        #region Methods
        /// <summary>
        /// Get the corresponding AbstractFastTransferStream.
        /// </summary>
        /// <returns>The corresponding AbstractFastTransferStream.</returns>
        public AbstractFastTransferStream GetAbstractFastTransferStream()
        {
            AbstractFastTransferStream afts = default(AbstractFastTransferStream);
            afts.StreamType = FastTransferStreamType.folderContent;

            afts.AbstractFolderContent.IsNoPermissionObjNotOut = !(this.SubFolders != null && this.SubFolders.Count > 0);
            afts.AbstractFolderContent.IsPidTagEcWarningOut = this.WarningCode != null;

            // The stack uses the structure defined in Open Specification 2.2.4.2 and the order to deserialize the payload. If the deserialization succeeds, the condition that IsFolderMessagesPrecededByPidTagFXDelProp is met.
            if (this.FolderMessages != null && this.folderMessages.FXDelPropList != null)
            {
                for (int i = 0; i < this.folderMessages.FXDelPropList.Count; i++)
                {
                    if ((this.folderMessages.FXDelPropList[i] == 0x3610000d) || (this.folderMessages.FXDelPropList[i] == 0x360F000d))
                    {
                        afts.AbstractFolderContent.AbsFolderMessage.IsFolderMessagesPrecededByPidTagFXDelProp = true;
                        break;
                    }
                }
            }

            if (this.SubFolders != null && this.SubFolders.Count > 0)
            {
                // afts.AbstractTopFolder.subFolderInScope = true;
                // The stack uses the structure defined in Open Specification 2.2.4.2 and the order to deserialize the payload. If the deserialization succeeds, the condition that IsFolderMessagesPrecededByPidTagFXDelProp is met.
                for (int i = 0; i < this.FXDelPropList.Count; i++)
                {
                    if (this.FXDelPropList[i] == 0x360E000d)
                    {
                        afts.AbstractFolderContent.IsSubFolderPrecededByPidTagFXDelProp = true;
                        break;
                    }
                }
            }

            // afts.AbstractFolderContent.AbsFolderMessage.MessageList.AbsMessage.AbsMessageContent.IsNoPermissionMessageNotOut = ((this.SubFolders == null) || (this.SubFolders != null && this.SubFolders.Count == 0));
            return afts;
        }

        /// <summary>
        ///  Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            // Actual stream deserialization:
            // folderContent        =  propList [PidTagEcWarning]
            //        ( (*(*PidTagFXDelProp PidTagNewFXFolder)) / folderMessages ) 
            //       [ *PidTagFXDelProp *SubFolder ]
            this.warningCode = null;
            this.fxdelPropList = new List<uint>();
            this.newFXFolderList = new List<Tuple<List<uint>, MS_OXCFXICS.FolderReplicaInfo>>();
            this.propList = new PropList(stream);
            if (!stream.IsEndOfStream)
            {
                uint marker = stream.VerifyUInt32();
                if (marker == (uint)MetaProperties.PidTagEcWarning)
                {
                    marker = stream.ReadUInt32();
                    this.warningCode = stream.ReadUInt32();
                    marker = stream.VerifyUInt32();
                }

                long lastPosi = stream.Position;
                while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp)
                    || stream.VerifyMetaProperty(MetaProperties.PidTagNewFXFolder))
                {
                    lastPosi = stream.Position;
                    List<uint> tempFXdel = new List<uint>();
                    while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp))
                    {
                        stream.ReadMetaProperty(MetaProperties.PidTagFXDelProp);
                        uint prop = stream.ReadUInt32();
                        tempFXdel.Add(prop);
                    }

                    if (!stream.IsEndOfStream)
                    {
                        marker = stream.VerifyUInt32();
                    }

                    if (marker == (uint)MetaProperties.PidTagNewFXFolder)
                    {
                        marker = stream.ReadUInt32();
                        stream.ReadUInt32();
                        FolderReplicaInfo fri = new FolderReplicaInfo(stream);
                        this.newFXFolderList.Add(
                            new Tuple<List<uint>, FolderReplicaInfo>(
                                tempFXdel, fri));
                    }
                    else
                    {
                        stream.Position = lastPosi;
                        marker = stream.VerifyUInt32();
                        break;
                    }
                }

                if (FolderMessages.Verify(stream))
                {
                    this.folderMessages = new FolderMessages(stream);
                }

                this.subFolders = new List<SubFolder>();
                while (stream.VerifyMetaProperty(MetaProperties.PidTagFXDelProp))
                {
                    stream.ReadMetaProperty(MetaProperties.PidTagFXDelProp);
                    uint prop = stream.ReadUInt32();
                    this.fxdelPropList.Add(prop);
                }

                while (SubFolder.Verify(stream))
                {
                    this.subFolders.Add(new SubFolder(stream));
                }
            }
        }
        #endregion
    }
}