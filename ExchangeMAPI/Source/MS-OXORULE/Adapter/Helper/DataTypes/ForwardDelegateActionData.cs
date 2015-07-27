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
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Action Data buffer format for ActionType: OP_FORWARD, OP_DELEGATE
    /// </summary>
    public class ForwardDelegateActionData : IActionData
    {
        /// <summary>
        /// Specifies recipient information
        /// </summary>
        private RecipientBlock[] recipientsData;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// Specifies the number of recipient blocks. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        private object recipientCount;

        /// <summary>
        /// Initializes a new instance of the ForwardDelegateActionData class.
        /// </summary>
        public ForwardDelegateActionData()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the ForwardDelegateActionData class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public ForwardDelegateActionData(CountByte countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets recipient information
        /// </summary>
        public RecipientBlock[] RecipientsData
        {
            get { return this.recipientsData; }
            set { this.recipientsData = value; }
        }

        /// <summary>
        /// Gets type of Count
        /// </summary>
        public CountByte CountType
        {
            get
            {
                return this.countType;
            }
        }

        /// <summary>
        /// Gets or sets the number of recipient blocks. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        public object RecipientCount
        {
            get
            {
                return this.recipientCount;
            }

            set
            {
                if (this.CountType == CountByte.TwoBytesCount)
                {
                    if (value is int)
                    {
                        this.recipientCount = (ushort)(int)value;
                    }
                    else
                    {
                        this.recipientCount = (ushort)value;
                    }
                }
                else
                {
                    if (value is int)
                    {
                        this.recipientCount = (uint)(int)value;
                    }
                    else
                    {
                        this.recipientCount = (uint)value;
                    }
                }
            }
        }

        /// <summary>
        /// The total Size of this ActionData buffer
        /// </summary>
        /// <returns>Number of bytes in this ActionData buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this ActionData
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> result = new List<byte>();
            if (this.CountType == CountByte.TwoBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((ushort)this.RecipientCount));
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((uint)this.RecipientCount));
            }

            // Add data of each RecipientBlock in RecipientsData
            foreach (RecipientBlock recipient in this.RecipientsData)
            {
                result.AddRange(recipient.Serialize());
            }

            return result.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to a ForwardActionData instance
        /// </summary>
        /// <param name="buffer">Byte array contains data of an ActionData instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            uint totalBytes = 0;
            uint count = 0;
            if (this.CountType == CountByte.TwoBytesCount)
            {
                this.RecipientCount = bufferReader.ReadUInt16();
                this.RecipientsData = new RecipientBlock[(ushort)this.RecipientCount];
                count = (uint)(ushort)this.RecipientCount;
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                this.RecipientCount = bufferReader.ReadUInt32();
                this.RecipientsData = new RecipientBlock[(uint)this.RecipientCount];
                count = (uint)this.RecipientCount;
            }

            totalBytes += bufferReader.Position;
            byte[] tempArray = bufferReader.ReadToEnd();
            uint bytesCount = 0;
            for (uint i = 0; i < count; i++)
            {
                bufferReader = new BufferReader(tempArray);
                tempArray = bufferReader.ReadBytes(bytesCount, (uint)(tempArray.Length - bytesCount));
                this.RecipientsData[i] = new RecipientBlock(this.CountType);
                bytesCount = this.RecipientsData[i].Deserialize(tempArray);
            }

            totalBytes += bufferReader.Position;
            return totalBytes;
        }
    }
}