namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;

    /// <summary>
    /// Message entry id structure.
    /// </summary>
    public class MessageEntryID
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
        /// Specify the message Type, 0x0007 means private message.
        /// </summary>
        private byte[] messageType = new byte[] { 0x07, 0x00 };

        /// <summary>
        /// A GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        private byte[] folderDataBaseGUID = new byte[16];

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder.
        /// </summary>
        private byte[] folderGlobalCounter = new byte[6];

        /// <summary>
        /// MUST be zero.
        /// </summary>
        private byte[] pad = new byte[] { 0x00, 0x00 };

        /// <summary>
        /// A GUID associated with the Store object of the message and corresponding to the DatabaseReplicationId field of the message ID.
        /// </summary>
        private byte[] messageDataBaseGUID = new byte[16];

        /// <summary>
        /// An unsigned 48-bit integer identifying the message.
        /// </summary>
        private byte[] messageGlobalCounter = new byte[6];

        /// <summary>
        /// Initializes a new instance of the MessageEntryID class.
        /// </summary>
        /// <param name="providerUID">Provider id value which can get in logon response.</param>
        /// <param name="folderDataBaseGUID">DatabaseGUID of the folder where the message is stored in.</param>
        /// <param name="folderGlobalCounter">GlobalCounter of the folder where the message is stored in.</param>
        /// <param name="messageDataBaseGUID">DatabaseGUID of the message.</param>
        /// <param name="messageGlobalCounter">GlboabCounter of message.</param>
        public MessageEntryID(byte[] providerUID, byte[] folderDataBaseGUID, byte[] folderGlobalCounter, byte[] messageDataBaseGUID, byte[] messageGlobalCounter)
        {
            this.providerUID = providerUID;
            this.folderDataBaseGUID = folderDataBaseGUID;
            this.folderGlobalCounter = folderGlobalCounter;
            this.messageDataBaseGUID = messageDataBaseGUID;
            this.messageGlobalCounter = messageGlobalCounter;
        }

        /// <summary>
        /// Gets or sets an unsigned 48-bit integer identifying the message.
        /// </summary>
        public byte[] MessageGlobalCounter
        {
            get { return this.messageGlobalCounter; }
            set { this.messageGlobalCounter = value; }
        }

        /// <summary>
        /// Gets or sets a GUID associated with the Store object of the message and corresponding to the DatabaseReplicationId field of the message ID.
        /// </summary>
        public byte[] MessageDataBaseGUID
        {
            get { return this.messageDataBaseGUID; }
            set { this.messageDataBaseGUID = value; }
        }

        /// <summary>
        /// Gets or sets an unsigned 48-bit integer identifying the folder.
        /// </summary>
        public byte[] FolderGlobalCounter
        {
            get { return this.folderGlobalCounter; }
            set { this.folderGlobalCounter = value; }
        }

        /// <summary>
        /// Gets or sets a GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        public byte[] FolderDataBaseGUID
        {
            get { return this.folderDataBaseGUID; }
            set { this.folderDataBaseGUID = value; }
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
        /// Gets or sets the message Type, 0x0007 means private message.
        /// </summary>
        public byte[] MessageType
        {
            get { return this.messageType; }
            set { this.messageType = value; }
        }

        /// <summary>
        /// Get folder entry id bytes array.
        /// </summary>
        /// <returns>Bytes array of folder entry id.</returns>
        public byte[] Serialize()
        {
            byte[] value = new byte[70];
            int index = 0;
            Array.Copy(this.flag, 0, value, index, this.flag.Length);
            index += this.flag.Length;
            Array.Copy(this.providerUID, 0, value, index, this.providerUID.Length);
            index += this.providerUID.Length;
            Array.Copy(this.messageType, 0, value, index, this.messageType.Length);
            index += this.messageType.Length;
            Array.Copy(this.folderDataBaseGUID, 0, value, index, this.folderDataBaseGUID.Length);
            index += this.folderDataBaseGUID.Length;
            Array.Copy(this.folderGlobalCounter, 0, value, index, this.folderGlobalCounter.Length);
            index += this.folderGlobalCounter.Length;
            Array.Copy(this.pad, 0, value, index, this.pad.Length);
            index += this.pad.Length;
            Array.Copy(this.messageDataBaseGUID, 0, value, index, this.messageDataBaseGUID.Length);
            index += this.messageDataBaseGUID.Length;
            Array.Copy(this.messageGlobalCounter, 0, value, index, this.messageGlobalCounter.Length);
            index += this.messageGlobalCounter.Length;
            Array.Copy(this.pad, 0, value, index, this.pad.Length);
            return value;
        }
    }
}