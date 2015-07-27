//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    /// <summary>
    /// The StringArray_r structure encodes an array of references to 8-bit character strings.
    /// </summary>
    public struct StringArray_r
    {
        /// <summary>
        /// The number of 8-bit character string references represented in the StringArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The 8-bit character string references. The strings referred to are NULL-terminated.
        /// </summary>
        [Size("CValues"), String(StringEncoding.ASCII)]
        public string[] LppzA;
    }

    /// <summary>
    /// The PropertyValue_r structure is an encoding of the PropertyValue data structure
    /// </summary>
    public struct PropertyValue_r
    {
        /// <summary>
        /// Encodes the proptag of the property whose value is represented by the PropertyValue_r data structure.
        /// </summary>
        public PropertyTag PropTag;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;

        /// <summary>
        /// Encodes the actual value of the property represented by the PropertyValue_r data structure.
        /// </summary>
        public byte[] Value;
    }

    /// <summary>
    /// The WStringArray_r structure encodes an array of references to Unicode strings.
    /// </summary>
    public struct WStringsArray_r
    {
        /// <summary>
        /// The number of Unicode character string references represented in the WStringArray_r structure. The value MUST NOT exceed 100,000.
        /// </summary>
        public uint CValues;

        /// <summary>
        /// The Unicode character string references. The strings are NULL-terminated.
        /// </summary>
        [Size("Count"), String(StringEncoding.Unicode)]
        public string[] LppszW;
    }

    /// <summary>
    /// The STAT structure is used to specify the state of a table and location information that applies to that table. 
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct STAT
    {
        /// <summary>
        /// This field contains a DWORD value that represents a sort order. 
        /// The client sets this field to specify the sort type of this table. 
        /// </summary>
        public uint SortType;

        /// <summary>
        /// This field contains a Minimal Entry ID. 
        /// The client sets this field to specify the Minimal Entry ID of the address book container that this STAT structure represents. 
        /// </summary>
        public uint ContainerID;

        /// <summary>
        /// This field contains a Minimal Entry ID. 
        /// The client sets this field to specify a beginning position in the table for the start of an NSPI method. 
        /// The server sets this field to report the end position in the table after processing an NSPI method.
        /// </summary>
        public uint CurrentRec;

        /// <summary>
        /// This field contains a long value. 
        /// The client sets this field to specify an offset from the beginning position in the table for the start of an NSPI method. 
        /// </summary>
        public int Delta;

        /// <summary>
        /// This field contains a DWORD value that specifies a position in the table.
        /// The client sets this field to specify a fractional position for the beginning position in the table for the start of an NSPI method.
        /// The server sets this field to specify the approximate fractional position at the end of an NSPI method. 
        /// </summary>
        public uint NumPos;

        /// <summary>
        /// This field contains a DWORD value that specifies the number of rows in the table. 
        /// The client sets this field to specify a fractional position for the beginning position in the table for the start of an NSPI method.
        /// The server sets this field to specify the total number of rows in the table. 
        /// </summary>
        public uint TotalRecs;

        /// <summary>
        /// This field contains a DWORD value that represents a code page. 
        /// The client sets this field to specify the code page the client uses for non-Unicode strings. 
        /// The server MUST use this value during string handling and MUST NOT modify this field.
        /// </summary>
        public uint CodePage;

        /// <summary>
        /// This field contains a DWORD value that represents a language code identifier (LCID).
        /// The client sets this field to specify the LCID associated with the template the client wants the server to return. 
        /// The server MUST NOT modify this field.
        /// </summary>
        public uint TemplateLocale;

        /// <summary>
        /// This field contains a DWORD value that represents an LCID. 
        /// The client sets this field to specify the LCID that it wants the server to use when sorting any strings. 
        /// The server MUST use this value during sorting and MUST NOT modify this field.
        /// </summary>
        public uint SortLocale;

        /// <summary>
        /// Parse the STAT from the response data.
        /// </summary>
        /// <param name="rawData">The response data.</param>
        /// <param name="startIndex">The start index.</param>
        /// <returns>The result of STAT.</returns>
        public static STAT Parse(byte[] rawData, ref int startIndex)
        {
            STAT state = new STAT();
            state.SortType = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.ContainerID = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.CurrentRec = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.Delta = BitConverter.ToInt32(rawData, startIndex);
            startIndex += sizeof(int);
            state.NumPos = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.TotalRecs = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.CodePage = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.TemplateLocale = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);
            state.SortLocale = BitConverter.ToUInt32(rawData, startIndex);
            startIndex += sizeof(uint);

            return state;
        }

        /// <summary>
        /// Initiate the stat to initial values.
        /// </summary>
        public void InitiateStat()
        {
            this.CodePage = (uint)RequiredCodePages.CP_TELETEX;
            this.CurrentRec = (uint)MinimalEntryIDs.MID_BEGINNING_OF_TABLE;
            this.ContainerID = 0;
            this.Delta = 0;
            this.NumPos = 0;
            this.SortLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            this.SortType = (uint)TableSortOrders.SortTypeDisplayName;
            this.TemplateLocale = (uint)DefaultLCID.NSPI_DEFAULT_LOCALE;
            this.TotalRecs = 0;
        }

        /// <summary>
        /// This method is used for serializing the STAT
        /// </summary>
        /// <returns>The serialized bytes of STAT</returns>
        public byte[] Serialize()
        {
            List<byte> buffer = new List<byte>();
            buffer.AddRange(BitConverter.GetBytes(this.SortType));
            buffer.AddRange(BitConverter.GetBytes(this.ContainerID));
            buffer.AddRange(BitConverter.GetBytes(this.CurrentRec));
            buffer.AddRange(BitConverter.GetBytes(this.Delta));
            buffer.AddRange(BitConverter.GetBytes(this.NumPos));
            buffer.AddRange(BitConverter.GetBytes(this.TotalRecs));
            buffer.AddRange(BitConverter.GetBytes(this.CodePage));
            buffer.AddRange(BitConverter.GetBytes(this.TemplateLocale));
            buffer.AddRange(BitConverter.GetBytes(this.SortLocale));

            return buffer.ToArray();
        }
    }

    /// <summary>
    /// The EphemeralEntryID structure identifies a specific object in the address book. 
    /// </summary>
    public struct EphemeralEntryID
    {
        /// <summary>
        /// The type of this ID. The value is the constant 0x87. 
        /// </summary>
        public byte IDType;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R1;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R2;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R3;

        /// <summary>
        /// Contain the GUID of the server that issues this Ephemeral Entry ID.
        /// </summary>
        public Guid ProviderUID;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000001.
        /// </summary>
        public uint R4;

        /// <summary>
        /// The display type of the object specified by this Ephemeral Entry ID. 
        /// </summary>
        public DisplayTypeValues DisplayType;

        /// <summary>
        /// The Minimal Entry ID of this object, as specified in section 2.3.8.1. 
        /// This value is expressed in little-endian format.
        /// </summary>
        public uint Mid;
    }

    /// <summary>
    /// The PermanentEntryID structure identifies a specific object in the address book. 
    /// </summary>
    public struct PermanentEntryID
    {
        /// <summary>
        /// The type of this ID. The value is the constant 0x00. 
        /// </summary>
        public byte IDType;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R1;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R2;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00.
        /// </summary>
        public byte R3;

        /// <summary>
        /// A FlatUID_r value that contains the constant GUID specified in Permanent Entry ID GUID.
        /// </summary>
        public Guid ProviderUID;

        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000001.
        /// </summary>
        public uint R4;

        /// <summary>
        /// The display type of the object specified by this Permanent Entry ID. 
        /// </summary>
        public DisplayTypeValues DisplayTypeString;

        /// <summary>
        /// The DN of the object specified by this Permanent Entry ID. The value is expressed as a DN.
        /// </summary>
        public string DistinguishedName;
    }
}