//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A folder entry id structure.
    /// </summary>
    public class ServerEID
    {
        /// <summary>
        /// MUST be 0x01.
        /// </summary>
        private byte ours = 0x01;

        /// <summary>
        /// For a folder in a private mailbox MUST be set to the MailboxGuid field value from the RopLogon
        /// </summary>
        private byte[] folderID = new byte[8];

        /// <summary>
        /// A GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        private byte[] messageID = new byte[8];

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder.
        /// </summary>
        private byte[] instance = new byte[4];

        /// <summary>
        /// Initializes a new instance of the ServerEID class.
        /// </summary>
        /// <param name="folderID">Identify store object is a mailbox or a public folder.</param>
        public ServerEID(byte[] folderID)
        {
            this.folderID = folderID;
            this.messageID = new byte[8];
            this.instance = new byte[4];
        }

        /// <summary>
        /// Gets or sets Flag that MUST be zero.
        /// </summary>
        public byte Ours
        {
            get { return this.ours; }
            set { this.ours = value; }
        }

        /// <summary>
        /// Gets or sets ProviderUID that for a folder in a private mailbox MUST be set to the MailboxGuid field value from the RopLogon
        /// </summary>
        public byte[] FolderID
        {
            get { return this.folderID; }
            set { this.folderID = value; }
        }

        /// <summary>
        /// Gets or sets a GUID associated with the Store object, and corresponding to the ReplicaId field of the FID.
        /// </summary>
        public byte[] MessageID
        {
            get { return this.messageID; }
            set { this.messageID = value; }
        }

        /// <summary>
        /// Gets or sets an unsigned 48-bit integer identifying the folder.
        /// </summary>
        public byte[] Instance
        {
            get { return this.instance; }
            set { this.instance = value; }
        }

        /// <summary>
        /// Get folder entry id bytes array.
        /// </summary>
        /// <returns>Bytes array of folder entry id.</returns>
        public byte[] Serialize()
        {
            byte[] value = new byte[21];
            int index = 0;
            value[0] = this.ours;
            index += 1;
            Array.Copy(this.folderID, 0, value, index, this.folderID.Length);
            index += this.folderID.Length;
            Array.Copy(this.messageID, 0, value, index, this.messageID.Length);
            index += this.messageID.Length;
            Array.Copy(this.instance, 0, value, index, this.instance.Length);
            return value;
        }

        /// <summary>
        /// Deserialize FolderEntryID.
        /// </summary>
        /// <param name="buffer">Entry id data array.</param>
        public void Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            byte currentOurs = reader.ReadByte();
            if (currentOurs != this.ours)
            {
                string errorMessage = "Wrong ours field, the expected value is " + this.ours.ToString() + ", the actual value is " + currentOurs.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            this.folderID = reader.ReadBytes(8);
            byte[] currentMessageID = reader.ReadBytes(8);
            if (!Common.CompareByteArray(currentMessageID, this.messageID))
            {
                string errorMessage = "Wrong messageID data, the expected value is " + this.messageID.ToString() + ", the actual value is " + currentMessageID.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte[] currentInstance = reader.ReadBytes(4);
            if (!Common.CompareByteArray(currentInstance, this.instance))
            {
                string errorMessage = "Wrong instance data, the expected value is " + this.instance.ToString() + ", the actual value is " + currentInstance.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }
        }
    }
}