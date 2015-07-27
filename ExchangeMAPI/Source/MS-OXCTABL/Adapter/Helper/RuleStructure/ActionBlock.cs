//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// ActionBlock structure in RuleAction
    /// </summary>
    public class ActionBlock
    {
        /// <summary>
        /// ActionData buffer, it is different base on different ActionType
        /// </summary>
        private IActionData actionDataValue;

        /// <summary>
        /// Specifies the types of action
        /// </summary>
        private ActionTypes actionType;

        /// <summary>
        /// MUST be used in conjunction with specific ActionTypes that support it, and MUST be zero otherwise
        /// </summary>
        private uint actionFlavor;

        /// <summary>
        /// Client-defined Flags. The ActionFlags field is used solely by the client, it is not used by the server but stored only
        /// </summary>
        private uint actionFlags;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private Count countType;

        /// <summary>
        /// MUST be the cumulative length (in BYTES) of the subsequent fields in this ActionBlock. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        private object actionLength;

        /// <summary>
        /// Initializes a new instance of the ActionBlock class.
        /// </summary>
        public ActionBlock()
        {
            this.countType = Count.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the ActionBlock class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class</param>
        public ActionBlock(Count countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets actionData buffer, it is different base on different ActionType
        /// </summary>
        public IActionData ActionDataValue
        {
            get { return this.actionDataValue; }
            set { this.actionDataValue = value; }
        }

        /// <summary>
        /// Gets or sets the types of action
        /// </summary>
        public ActionTypes ActionType
        {
            get { return this.actionType; }
            set { this.actionType = value; }
        }

        /// <summary>
        /// Gets or sets the ActionFlavor
        /// </summary>
        public uint ActionFlavor
        {
            get { return this.actionFlavor; }
            set { this.actionFlavor = value; }
        }

        /// <summary>
        /// Gets or sets Client-defined Flags.
        /// </summary>
        public uint ActionFlags
        {
            get { return this.actionFlags; }
            set { this.actionFlags = value; }
        }

        /// <summary>
        /// Gets type of Count
        /// </summary>
        public Count CountType
        {
            get
            {
                return this.countType;
            }
        }

        /// <summary>
        /// Gets or sets the action length that MUST be the cumulative length (in BYTES) of the subsequent fields in this ActionBlock. If the countType is FourBytesCount, its type is uint. Otherwise, its Type is ushort.
        /// </summary>
        public object ActionLength
        {
            get
            {
                return this.actionLength;
            }

            set
            {
                if (this.CountType == Count.TwoBytesCount)
                {
                    if (value is int)
                    {
                        this.actionLength = (ushort)(int)value;
                    }
                    else
                    {
                        this.actionLength = (ushort)value;
                    }
                }
                else
                {
                    if (value is int)
                    {
                        this.actionLength = (uint)(int)value;
                    }
                    else
                    {
                        this.actionLength = (uint)value;
                    }
                }
            }
        }

        /// <summary>
        /// Get the total Size of ActionData
        /// </summary>
        /// <returns>The Size of ActionData buffer</returns>
        public int Size()
        {
            int lengthOfActionLength = 0;
            if (this.CountType == Count.TwoBytesCount)
            {
                lengthOfActionLength += 2;
            }
            else if (this.CountType == Count.FourBytesCount)
            {
                lengthOfActionLength += 4;
            }

            // Length of ActionType is 1
            // Length of ActionFlavor is 4
            // Length of ActionFlags is 4
            return lengthOfActionLength + 1 + 4 + 4 + this.ActionDataValue.Size();
        }

        /// <summary>
        /// Get serialized byte array for this structure
        /// </summary>
        /// <returns>Serialized byte array</returns>
        public byte[] Serialize()
        {
            List<byte> result = new List<byte>();
            if (this.CountType == Count.TwoBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((ushort)this.ActionLength));
            }
            else if (this.CountType == Count.FourBytesCount)
            {
                result.AddRange(BitConverter.GetBytes((uint)this.ActionLength));
            }

            result.Add((byte)this.ActionType);
            result.AddRange(BitConverter.GetBytes(this.ActionFlavor));
            result.AddRange(BitConverter.GetBytes(this.ActionFlags));
            if (this.ActionDataValue.Serialize() != null)
            {
                result.AddRange(this.ActionDataValue.Serialize());
            }

            return result.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to an ActionBlock instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an ActionBlock instance</param>
        /// <returns>Bytes count that deserialized in buffer</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            uint totalBytes = 0;
            if (this.CountType == Count.TwoBytesCount)
            {
                this.ActionLength = bufferReader.ReadUInt16();
                totalBytes += bufferReader.Position;
                bufferReader = new BufferReader(bufferReader.ReadBytes((ushort)this.ActionLength));
            }
            else if (this.CountType == Count.FourBytesCount)
            {
                this.ActionLength = bufferReader.ReadUInt32();
                totalBytes += bufferReader.Position;
                bufferReader = new BufferReader(bufferReader.ReadBytes((uint)this.ActionLength));
            }

            this.ActionType = (ActionTypes)bufferReader.ReadByte();
            this.ActionFlavor = bufferReader.ReadUInt32();
            this.ActionFlags = bufferReader.ReadUInt32();

            totalBytes += bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();
            this.ActionDataValue = new DeleteMarkReadActionData();
            totalBytes += this.ActionDataValue.Deserialize(tmpArray);
            return totalBytes;
        }
    }
}