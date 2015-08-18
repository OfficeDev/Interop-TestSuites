namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A folder entry id structure.
    /// </summary>
    public class FolderEntryID
    {
        /// <summary>
        /// MUST be zero.
        /// </summary>
        private byte[] flag = new byte[] { 0x00, 0x00, 0x00, 0x00 };

        /// <summary>
        /// For a folder in a private mailbox MUST be set to the MailboxGuid field value from the RopLogon
        /// </summary>
        private byte[] providerUID = new byte[16];

        /// <summary>
        /// Specify the folder type, 0x0001 means private folder.
        /// </summary>
        private byte[] folderType = new byte[2];

        /// <summary>
        /// A GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        private byte[] dataBaseGUID = new byte[16];

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder.
        /// </summary>
        private byte[] globalCounter = new byte[6];

        /// <summary>
        /// MUST be zero.
        /// </summary>
        private byte[] pad = new byte[] { 0x00, 0x00 };

        /// <summary>
        /// Initializes a new instance of the FolderEntryID class.
        /// </summary>
        /// <param name="objectType">Identify store object is a mailbox or a public folder.</param>
        /// <param name="providerUID">Provider id value which can get in logon response.</param>
        /// <param name="databaseGUID">DatabaseGUID of specific folder, which can get in folder's longterm id.</param>
        /// <param name="globalCounter">Global counter of specific folder, which can get in folder's longterm id.</param>
        public FolderEntryID(StoreObjectType objectType, byte[] providerUID, byte[] databaseGUID, byte[] globalCounter)
        {
            this.providerUID = providerUID;
            this.dataBaseGUID = databaseGUID;
            this.globalCounter = globalCounter;
            if (objectType == StoreObjectType.Mailbox)
            {
                this.folderType = new byte[] { 0x01, 0x00 };
            }
            else
            {
                this.folderType = new byte[] { 0x03, 0x00 };
            }
        }

        /// <summary>
        /// Initializes a new instance of the FolderEntryID class.
        /// </summary>
        /// <param name="objectType">Identify store object is a mailbox or a public folder.</param>
        public FolderEntryID(StoreObjectType objectType)
        {
            if (objectType == StoreObjectType.Mailbox)
            {
                this.folderType = new byte[] { 0x01, 0x00 };
            }
            else
            {
                this.folderType = new byte[] { 0x03, 0x00 };
            }
        }

        /// <summary>
        /// Gets or sets an unsigned 48-bit integer identifying the folder.
        /// </summary>
        public byte[] GlobalCounter
        {
            get { return this.globalCounter; }
            set { this.globalCounter = value; }
        }

        /// <summary>
        /// Gets or sets a GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        public byte[] DataBaseGUID
        {
            get { return this.dataBaseGUID; }
            set { this.dataBaseGUID = value; }
        }

        /// <summary>
        /// Gets or sets ProviderUID that for a folder in a private mailbox MUST be set to the MailboxGuid field value from the RopLogon
        /// </summary>
        public byte[] ProviderUID
        {
            get { return this.providerUID; }
            set { this.providerUID = value; }
        }

        /// <summary>
        /// Gets or sets Pad that MUST be zero.
        /// </summary>
        public byte[] Pad
        {
            get { return this.pad; }
            set { this.pad = value; }
        }

        /// <summary>
        /// Gets or sets Flag that MUST be zero.
        /// </summary>
        public byte[] Flag
        {
            get { return this.flag; }
            set { this.flag = value; }
        }

        /// <summary>
        /// Gets or sets the folder type, 0x0001 means private folder.
        /// </summary>
        public byte[] FolderType
        {
            get { return this.folderType; }
            set { this.folderType = value; }
        }

        /// <summary>
        /// Get folder entry id bytes array.
        /// </summary>
        /// <returns>Bytes array of folder entry id.</returns>
        public byte[] Serialize()
        {
            byte[] value = new byte[46];
            int index = 0;
            Array.Copy(this.flag, 0, value, index, this.flag.Length);
            index += this.flag.Length;
            Array.Copy(this.providerUID, 0, value, index, this.providerUID.Length);
            index += this.providerUID.Length;
            Array.Copy(this.folderType, 0, value, index, this.folderType.Length);
            index += this.folderType.Length;
            Array.Copy(this.dataBaseGUID, 0, value, index, this.dataBaseGUID.Length);
            index += this.dataBaseGUID.Length;
            Array.Copy(this.globalCounter, 0, value, index, this.globalCounter.Length);
            index += this.globalCounter.Length;
            Array.Copy(this.pad, 0, value, index, this.pad.Length);
            return value;
        }

        /// <summary>
        /// Deserialize FolderEntryID.
        /// </summary>
        /// <param name="buffer">Entry id data array.</param>
        public void Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            byte[] currentFlag = reader.ReadBytes(4);
            if (!Common.CompareByteArray(currentFlag, this.flag))
            {
                if (currentFlag != null)
                {
                    string errorMessage = "Wrong flag error, the expect flag is { 0x00, 0x00, 0x00, 0x00}, actual is " + currentFlag.ToString() + "!";
                    throw new ArgumentException(errorMessage);
                }
                else
                {
                    string errorMessage = "Wrong flag error, the expect flag is { 0x00, 0x00, 0x00, 0x00}, actual is null!";
                    throw new ArgumentException(errorMessage);
                }
            }

            this.providerUID = reader.ReadBytes(16);
            byte[] currentFolderType = reader.ReadBytes(2);
            if (!Common.CompareByteArray(currentFolderType, this.folderType))
            {
                if (currentFlag != null)
                {
                    string errorMessage = "Wrong folder type error, the expect folder type is " + this.folderType.ToString() + ", the actual is " + currentFolderType.ToString() + "!";
                    throw new ArgumentException(errorMessage);
                }
                else
                {
                    string errorMessage = "Wrong folder type error, the expect folder type is " + this.folderType.ToString() + ", actual is null!";
                    throw new ArgumentException(errorMessage);
                }
            }

            this.dataBaseGUID = reader.ReadBytes(16);
            this.globalCounter = reader.ReadBytes(6);
            byte[] currentPad = reader.ReadBytes(2);
            if (!Common.CompareByteArray(currentPad, this.pad))
            {
                if (currentFlag != null)
                {
                    string errorMessage = "Wrong pad data, the expect pad data is " + this.pad.ToString() + ", the actual is " + currentPad.ToString() + "!";
                    throw new Exception(errorMessage);
                }
                else
                {
                    string errorMessage = "Wrong pad data, the expect pad data is " + this.pad.ToString() + ", actual is null!";
                    throw new Exception(errorMessage);
                }
            }
        }
    }
}