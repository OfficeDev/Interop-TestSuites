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
    /// For Rule Action format
    /// </summary>
    public class RuleAction
    {
        /// <summary>
        /// Specifies Actions in the Rule
        /// </summary>
        private ActionBlock[] actions;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// Specifies the number of ActionBlocks that are packed in this buffer. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        private object actionsNumber;

        /// <summary>
        /// Initializes a new instance of the RuleAction class.
        /// </summary>
        public RuleAction()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the RuleAction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public RuleAction(CountByte countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets Actions in the Rule 
        /// </summary>
        public ActionBlock[] Actions
        {
            get { return this.actions; }
            set { this.actions = value; }
        }

        /// <summary>
        /// Gets type of Count
        /// </summary>
        public CountByte CountType
        {
            get { return this.countType; }
        }

        /// <summary>
        /// Gets or sets the number of ActionBlocks that are packed in this buffer. This number MUST be greater than zero. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        public object NoOfActions
        {
            get
            {
                return this.actionsNumber;
            }

            set
            {
                if (this.CountType == CountByte.TwoBytesCount)
                {
                    if (value is int)
                    {
                        this.actionsNumber = (ushort)(int)value;
                    }
                    else
                    {
                        this.actionsNumber = (ushort)value;
                    }
                }
                else
                {
                    if (value is int)
                    {
                        this.actionsNumber = (uint)(int)value;
                    }
                    else
                    {
                        this.actionsNumber = (uint)value;
                    }
                }
            }
        }

        /// <summary>
        /// Get the total Size of ActionData
        /// </summary>
        /// <returns>The Size of RuleAction buffer.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public byte[] Serialize()
        {
            List<byte> result = new List<byte>();
            if (this.CountType == CountByte.TwoBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((ushort)this.NoOfActions));
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((uint)this.NoOfActions));
            }

            foreach (ActionBlock actionBlock in this.Actions)
            {
                result.AddRange(actionBlock.Serialize());
            }

            return result.ToArray();
        }

        /// <summary>
        /// Parse method to obtain current structure from byte array
        /// </summary>
        /// <param name="buffer">Byte array data.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            uint count = 0;
            if (this.CountType == CountByte.TwoBytesCount)
            {
                this.NoOfActions = bufferReader.ReadUInt16();
                this.Actions = new ActionBlock[(ushort)this.NoOfActions];
                count = (uint)(ushort)this.NoOfActions;
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                this.NoOfActions = bufferReader.ReadUInt32();
                this.Actions = new ActionBlock[(uint)this.NoOfActions];
                count = (uint)this.NoOfActions;
            }

            uint totalBytes = bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();
            uint bytesCount = 0;
            for (uint i = 0; i < count; i++)
            {
                bufferReader = new BufferReader(tmpArray);
                tmpArray = bufferReader.ReadBytes(bytesCount, (uint)(tmpArray.Length - bytesCount));
                this.Actions[i] = new ActionBlock(this.CountType);
                bytesCount = this.Actions[i].Deserialize(tmpArray);
                totalBytes += bytesCount;
            }

            return totalBytes;
        }
    }
}