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
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// multi-valued property type PropValue
    /// </summary>
    public class MvPropTypePropValue : PropValue
    {
        /// <summary>
        /// This represent the length variable.
        /// </summary>
        private int length;

        /// <summary>
        /// A list of  serialized fixed size values.
        /// </summary>
        private List<byte[]> fixedSizeValueList;

        /// <summary>
        /// A list of variate size value and value's length 's TUPLEs
        /// </summary>
        private List<Tuple<int, byte[]>> varSizeValueList;

        /// <summary>
        /// Initializes a new instance of the MvPropTypePropValue class.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public MvPropTypePropValue(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the length.
        /// </summary>
        public int Length
        {
            get { return this.length; }
            set { this.length = value; }
        }

        /// <summary>
        /// Gets or sets the fixedSizeValueList.
        /// </summary>
        public List<byte[]> FixedSizeValue
        {
            get { return this.fixedSizeValueList; }
            set { this.fixedSizeValueList = value; }
        }

        /// <summary>
        /// Gets or sets the varSizeValue list.
        /// </summary>
        public List<Tuple<int, byte[]>> VarSizeValueList
        {
            get { return this.varSizeValueList; }
            set { this.varSizeValueList = value; }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized MvPropTypePropValue.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>I the stream's current position contains 
        /// a serialized MvPropTypePropValue, return true, else false</returns>
        public static new bool Verify(FastTransferStream stream)
        {
            ushort tmp = stream.VerifyUInt16();
            return LexicalTypeHelper.IsMVType((PropertyDataType)tmp) && !PropValue.IsPidTagIdsetGiven(stream);
        }

        /// <summary>
        /// Deserialize a MvPropTypePropValue instance from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <returns>A MvPropTypePropValue instance </returns>
        public static new LexicalBase DeserializeFrom(FastTransferStream stream)
        {
            return new MvPropTypePropValue(stream);
        }

        /// <summary>
        /// Deserialize next object from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        public override void ConsumeNext(FastTransferStream stream)
        {
            base.ConsumeNext(stream);
            PropertyDataType type = (PropertyDataType)this.PropType;
            this.length = stream.ReadInt16();
            switch (type)
            {
                case PropertyDataType.PtypMultipleInteger16:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 2);
                    break;
                case PropertyDataType.PtypMultipleInteger32:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 4);
                    break;
                case PropertyDataType.PtypMultipleFloating32:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 4);
                    break;
                case PropertyDataType.PtypMultipleFloating64:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 8);
                    break;
                case PropertyDataType.PtypMultipleCurrency:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 8);
                    break;
                case PropertyDataType.PtypMultipleFloatingTime:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 8);
                    break;
                case PropertyDataType.PtypMultipleInteger64:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 8);
                    break;
                case PropertyDataType.PtypMultipleTime:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, 8);
                    break;
                case PropertyDataType.PtypMultipleGuid:
                    this.fixedSizeValueList = stream.ReadBlocks(this.length, Guid.Empty.ToByteArray().Length);
                    break;
                case PropertyDataType.PtypMultipleBinary:
                    this.varSizeValueList = stream.ReadLengthBlocks(this.length);
                    break;
                case PropertyDataType.PtypMultipleString:
                    this.varSizeValueList = stream.ReadLengthBlocks(this.length);
                    break;
                case PropertyDataType.PtypMultipleString8:
                    this.varSizeValueList = stream.ReadLengthBlocks(this.length);
                    break;
            }
        }
    }
}