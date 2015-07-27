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
    /// ActionBlock struct in RuleAction
    /// </summary>
    public class ActionBlock
    {
        /// <summary>
        /// ActionData buffer. It is different base on different ActionType
        /// </summary>
        private IActionData actionDataValue;

        /// <summary>
        /// Specifies the types of action
        /// </summary>
        private ActionType actionType;

        /// <summary>
        /// MUST be used in conjunction with specific ActionType that support it, and MUST be zero otherwise
        /// </summary>
        private uint actionFlavor;

        /// <summary>
        /// Client-defined Flags. The ActionFlags field is used solely by the client, and it is not used by the server but stored only
        /// </summary>
        private uint actionFlags;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// MUST be the cumulative length (in BYTES) of the subsequent fields in this ActionBlock, this Type is uint, otherwise, its Type is ushort.
        /// </summary>
        private object actionLength;

        /// <summary>
        /// Initializes a new instance of the ActionBlock class.
        /// </summary>
        public ActionBlock()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the ActionBlock class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public ActionBlock(CountByte countType)
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
        public ActionType ActionType
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
        public CountByte CountType
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
                if (this.CountType == CountByte.TwoBytesCount)
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
        /// <returns>The Size of ActionData buffer.</returns>
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
                result.AddRange(BitConverter.GetBytes((ushort)this.ActionLength));
            }
            else if (this.CountType == CountByte.FourBytesCount)
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
        /// <param name="buffer">Byte array contain data of an ActionBlock instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader bufferReader = new BufferReader(buffer);
            uint totalBytes = 0;
            if (this.CountType == CountByte.TwoBytesCount)
            {
                this.ActionLength = bufferReader.ReadUInt16();
                totalBytes += bufferReader.Position;
                bufferReader = new BufferReader(bufferReader.ReadBytes((ushort)this.ActionLength));
            }
            else if (this.CountType == CountByte.FourBytesCount)
            {
                this.ActionLength = bufferReader.ReadUInt32();
                totalBytes += bufferReader.Position;
                bufferReader = new BufferReader(bufferReader.ReadBytes((uint)this.ActionLength));
            }

            this.ActionType = (ActionType)bufferReader.ReadByte();
            this.ActionFlavor = bufferReader.ReadUInt32();
            this.ActionFlags = bufferReader.ReadUInt32();

            totalBytes += bufferReader.Position;
            byte[] tmpArray = null;
            byte[] remainBuffer = bufferReader.ReadToEnd();
            tmpArray = remainBuffer;

            switch (this.ActionType)
            {
                case ActionType.OP_MOVE:
                case ActionType.OP_COPY:

                    // On Exchange 2013, a redundant "0xff" field may be inserted before the actual Action data. 
                    if (remainBuffer != null && remainBuffer[0] == 0xff)
                    {
                        tmpArray = new byte[remainBuffer.Length - 1];
                        Array.Copy(remainBuffer, 1, tmpArray, 0, remainBuffer.Length - 1);
                        this.ActionLength = (ushort)this.ActionLength - 1;
                    }

                    if (this.CountType == CountByte.TwoBytesCount)
                    {
                        this.ActionDataValue = new MoveCopyActionData();
                    }
                    else if (this.CountType == CountByte.FourBytesCount)
                    {
                        this.ActionDataValue = new MoveCopyActionDataOfExtendedRule();
                    }

                    break;
                case ActionType.OP_REPLY:
                case ActionType.OP_OOF_REPLY:

                    if (this.CountType == CountByte.TwoBytesCount)
                    {
                        this.ActionDataValue = new ReplyActionData();
                    }
                    else if (this.CountType == CountByte.FourBytesCount)
                    {
                        this.ActionDataValue = new ReplyActionDataOfExtendedRule();
                    }

                    break;
                case ActionType.OP_DEFER_ACTION:
                    this.ActionDataValue = new DeferredActionData();
                    break;
                case ActionType.OP_BOUNCE:
                    this.ActionDataValue = new BounceActionData();
                    break;
                case ActionType.OP_FORWARD:
                case ActionType.OP_DELEGATE:
                    this.ActionDataValue = new ForwardDelegateActionData(this.CountType);
                    break;
                case ActionType.OP_TAG:
                    this.ActionDataValue = new TagActionData();
                    break;
                case ActionType.OP_DELETE:
                case ActionType.OP_MARK_AS_READ:
                    this.ActionDataValue = new DeleteMarkReadActionData();
                    break;
            }

            totalBytes += this.ActionDataValue.Deserialize(tmpArray);
            return totalBytes;
        }
    }
}