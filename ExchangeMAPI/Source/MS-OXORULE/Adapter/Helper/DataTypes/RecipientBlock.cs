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
    /// Recipient Data buffer structure which used for OP_FORWARD, OP_DELEGATE ActionType
    /// </summary>
    public class RecipientBlock
    {
        /// <summary>
        /// Must be set to 0x01
        /// </summary>
        private byte reserved = 0x01;

        /// <summary>
        /// Specifies the properties values in this data buffer
        /// </summary>
        private TaggedPropertyValue[] propertiesData;

        /// <summary>
        /// Specifies the number of properties in the block. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        private object propertiesNumber;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// Initializes a new instance of the RecipientBlock class.
        /// </summary>
        public RecipientBlock()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the RecipientBlock class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public RecipientBlock(CountByte countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets the Reserved
        /// </summary>
        public byte Reserved
        {
            get { return this.reserved; }
            set { this.reserved = value; }
        }

        /// <summary>
        /// Gets or sets the properties values in this data buffer
        /// </summary>
        public TaggedPropertyValue[] PropertiesData
        {
            get { return this.propertiesData; }
            set { this.propertiesData = value; }
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
        /// Gets or sets the value that specifies the number of properties in the block. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        public object NoOfProperties
        {
            get
            {
                return this.propertiesNumber;
            }

            set
            {
                if (this.CountType == CountByte.TwoBytesCount)
                {
                    if (value is int)
                    {
                        this.propertiesNumber = (ushort)(int)value;
                    }
                    else
                    {
                        this.propertiesNumber = (ushort)value;
                    }
                }
                else
                {
                    if (value is int)
                    {
                        this.propertiesNumber = (uint)(int)value;
                    }
                    else
                    {
                        this.propertiesNumber = (uint)value;
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
            List<byte> result = new List<byte>
            {
                this.Reserved
            };
            if (this.CountType == CountByte.TwoBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((ushort)this.NoOfProperties));
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((uint)this.NoOfProperties));
            }

            // Add data of each property value
            foreach (TaggedPropertyValue taggedPValue in this.PropertiesData)
            {
                result.AddRange(taggedPValue.Serialize());
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
            this.Reserved = bufferReader.ReadByte();

            uint count = 0;
            if (this.CountType == CountByte.TwoBytesCount)
            {
                this.NoOfProperties = bufferReader.ReadUInt16();
                this.PropertiesData = new TaggedPropertyValue[(ushort)this.NoOfProperties];
                count = (uint)(ushort)this.NoOfProperties;
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                this.NoOfProperties = bufferReader.ReadUInt32();
                this.PropertiesData = new TaggedPropertyValue[(uint)this.NoOfProperties];
                count = (uint)this.NoOfProperties;
            }

            uint size = bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();
            for (uint i = 0; i < count; i++)
            {
                TaggedPropertyValue tagValue = AdapterHelper.ReadTaggedProperty(tmpArray);
                this.PropertiesData[i] = tagValue;
                uint tagSize = (uint)tagValue.Size();
                size += tagSize;

                bufferReader = new BufferReader(tmpArray);
                tmpArray = bufferReader.ReadBytes(tagSize, (uint)(tmpArray.Length - tagSize));
            }

            return size;
        }
    }
}