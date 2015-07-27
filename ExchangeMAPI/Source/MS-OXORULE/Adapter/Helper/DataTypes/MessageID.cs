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

    /// <summary>
    /// MessageId structure.
    /// </summary>
    public class MessageID
    {
        /// <summary>
        /// A 16-bit unsigned integer identifying a Store object.
        /// </summary>
        private byte[] replicaId;

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder within its Store object.
        /// </summary>
        private byte[] globalCounter;

        /// <summary>
        /// Size of message id.
        /// </summary>
        private int size;

        /// <summary>
        /// Initializes a new instance of the MessageID class.
        /// </summary>
        public MessageID()
        {
            this.ReplicaId = new byte[2];
            this.GlobalCounter = new byte[6];
            this.size = 0;
        }

        /// <summary>
        /// Gets or sets a 16-bit unsigned integer identifying a Store object.
        /// </summary>
        public byte[] ReplicaId
        {
            get { return this.replicaId; }
            set { this.replicaId = value; }
        }

        /// <summary>
        /// Gets or sets an unsigned 48-bit integer identifying the folder within its Store object.
        /// </summary>
        public byte[] GlobalCounter
        {
            get { return this.globalCounter; }
            set { this.globalCounter = value; }
        }

        /// <summary>
        /// Gets or sets size of message id.
        /// </summary>
        public int Size
        {
            get { return this.size; }
            set { this.size = value; }
        }

        /// <summary>
        /// Deserialize folderId structure.
        /// </summary>
        /// <param name="folderId">Folder id of the folder object.</param>
        public void Deserialize(ulong folderId)
        {
            byte[] folderIdBytes = BitConverter.GetBytes(folderId);
            int index = 0;
            Array.Copy(folderIdBytes, 0, this.ReplicaId, 0, this.ReplicaId.Length);
            index += this.ReplicaId.Length;
            Array.Copy(folderIdBytes, index, this.GlobalCounter, 0, this.GlobalCounter.Length);
            this.size = this.ReplicaId.Length + this.GlobalCounter.Length;
        }
    }
}