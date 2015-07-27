//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// The FolderChange element contains a new or changed folder in the
    /// hierarchy synchronization.
    /// folderChange         = IncrSyncChg propList
    /// </summary>
    public class FolderChange : SyntacticalBase
    {
        /// <summary>
        /// A propList value.
        /// </summary>
        private PropList propList;

        /// <summary>
        /// Initializes a new instance of the FolderChange class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public FolderChange(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the propList.
        /// </summary>
        public PropList PropList
        {
            get { return this.propList; }
            set { this.propList = value; }
        }

        /// <summary>
        /// Gets a value indicating whether PidTagFolderId property is existent;
        /// </summary>
        public bool HasPidTagFolderId
        {
            get
            {
                if (this.PropList != null)
                {
                    return this.PropList.HasPropertyTag(0x6748, 0x0014);
                }

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether PidTagParentFolderId  property is existent.
        /// </summary>
        public bool HasPidTagParentFolderId
        {
            get
            {
                if (this.PropList != null)
                {
                    return this.PropList.HasPropertyTag(0x6749, 0x0014);
                }

                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether PidTagParentSourceKey is zero.
        /// </summary>
        public bool IsPidTagParentSourceKeyZero
        {
            get
            {
                byte[] buffer = this.PropList.GetPropValue(0x65e1, 0x0102) as byte[];
                for (int i = 0; i < buffer.Length; i++)
                {
                    if (buffer[i] != 0)
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        /// <summary>
        /// Gets a value indicating whether PidTagSourceKey is zero.
        /// </summary>
        public bool IsPidTagSourceKeyZero
        {
            get
            {
                byte[] buffer = this.PropList.GetPropValue(0x65e0, 0x0102) as byte[];
                for (int i = 0; i < buffer.Length; i++)
                {
                    if (buffer[i] != 0)
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized folderChange.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized folderChange, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return stream.VerifyMarker(Markers.PidTagIncrSyncChg);
        }

        /// <summary>
        /// Get the corresponding abstractFolderChange.
        /// </summary>
        /// <returns>The corresponding abstractFolderChange.</returns>
        public AbstractFolderChange GetAbstractFolderChange()
        {
            AbstractFolderChange fc = default(AbstractFolderChange);
            fc.IsPidTagFolderIdExist = this.HasPidTagFolderId;
            fc.IsPidTagParentFolderIdExist = this.HasPidTagParentFolderId;
            fc.IsPidTagParentSourceKeyValueZero = this.IsPidTagParentSourceKeyZero;
            fc.IsPidTagSourceKeyValueZero = this.IsPidTagSourceKeyZero;
            return fc;
        }

        /// <summary>
        /// Get parent source key bytes.
        /// </summary>
        /// <returns>A byte array.</returns>
        public byte[] GetParentSourceKey()
        {
            return this.PropList.GetPropValue(0x65e1, 0x0102) as byte[];
        }

        /// <summary>
        /// Gets source key bytes.
        /// </summary>
        /// <returns>A byte array.</returns>
        public byte[] GetSourceKey()
        {
            return this.PropList.GetPropValue(0x65e0, 0x0102) as byte[];
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            if (stream.ReadMarker(Markers.PidTagIncrSyncChg))
            {
                this.propList = new PropList(stream);
                return;
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
        }
    }
}