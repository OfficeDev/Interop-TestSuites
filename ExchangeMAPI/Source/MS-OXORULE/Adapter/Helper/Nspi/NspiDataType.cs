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
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    /// <summary>
    /// These values are used to specify Instruction types.
    /// </summary>
    public enum InstructionTypeValues : uint
    {
        /// <summary>
        /// Halt Instruction.
        /// </summary>
        Halt = 0x00000000,
        
        /// <summary>
        /// Error Instruction.
        /// </summary>
        Error = 0x00000001,
        
        /// <summary>
        /// Emit Property Value Instruction.
        /// </summary>
        Emit_Property_Value = 0x00000002,
        
        /// <summary>
        /// Jump Instruction.
        /// </summary>
        Jump = 0x00000003,
        
        /// <summary>
        /// Jump If Not Exists Instruction.
        /// </summary>
        Jump_If_Not_Exists = 0x00000004,
        
        /// <summary>
        /// Jump If Equal Properties Instruction.
        /// </summary>
        Jump_If_Equal_Properties = 0x00000005,
        
        /// <summary>
        /// Emit Upper Property Instruction.
        /// </summary>
        Emit_Upper_Property = 0x00000006,
        
        /// <summary>
        /// Emit string Instruction.
        /// </summary>
        Emit_String = 0x80000002,
        
        /// <summary>
        /// Jump If Equal Values Instruction.
        /// </summary>
        Jump_If_Equal_Values = 0x40000005,
        
        /// <summary>
        /// Emit Upper string Instruction.
        /// </summary>
        Emit_Upper_String = 0x80000006
    }

    /// <summary>
    /// These values are used to specify display types.
    /// </summary>
    public enum DisplayTypeValues : uint
    {
        /// <summary>
        /// A typical messaging user.
        /// </summary>
        DT_MAILUSER = 0x00000000,

        /// <summary>
        /// A distribution list.
        /// </summary>
        DT_DISTLIST = 0x00000001,

        /// <summary>
        /// A forum, such as a bulletin board service or a public or shared folder.
        /// </summary>
        DT_FORUM = 0x00000002,

        /// <summary>
        /// An automated agent, such as Quote-Of-The-Day or a weather chart display.
        /// </summary>
        DT_AGENT = 0x00000003,

        /// <summary>
        /// An address book object defined for a large group, such as helpdesk, accounting, coordinator, or department.
        /// </summary>
        DT_ORGANIZATION = 0x00000004,

        /// <summary>
        /// A private, personally administered distribution list.
        /// </summary>
        DT_PRIVATE_DISTLIST = 0x00000005,

        /// <summary>
        /// An address book object known to be from a foreign or remote messaging system.
        /// </summary>
        DT_REMOTE_MAILUSER = 0x00000006,

        /// <summary>
        /// An address book hierarchy table container.
        /// </summary>
        DT_CONTAINER = 0x00000100,

        /// <summary>
        /// A Display Template object.
        /// </summary>
        DT_TEMPLATE = 0x00000101,

        /// <summary>
        /// An Address Creation Template.
        /// </summary>
        DT_ADDRESS_TEMPLATE = 0x00000102,

        /// <summary>
        /// A Search Template.
        /// </summary>
        DT_SEARCH = 0x00000200
    }

    /// <summary>
    /// These values appear as bit Flags in the NspiGetTemplateInfo method. 
    /// </summary>
    public enum TemplateInfoFlags : uint
    {
        /// <summary>
        /// Specifies that the server is to return the value that represents a template.
        /// </summary>
        TI_TEMPLATE = 0x00000001,

        /// <summary>
        /// Specifies that the server is to return the value of the script associated with a template.
        /// </summary>
        TI_SCRIPT = 0x00000004,

        /// <summary>
        /// Specifies that the server is to return the e-mail Type associated with a template.
        /// </summary>
        TI_EMT = 0x00000010,

        /// <summary>
        /// Specifies that the server is to return the name of the help file associated with a template.
        /// </summary>
        TI_HELPFILE_NAME = 0x00000020,

        /// <summary>
        /// Specifies that the server is to return the contents of the help file associated with a template.
        /// </summary>
        TI_HELPFILE_CONTENTS = 0x00000040,
    }
    
    /// <summary>
    /// These values appear as bit Flags in theNspiGetSpecialTable method.
    /// </summary>
    public enum SpecialTableFlags : uint
    {
        /// <summary>
        /// Specifies that the NSPI server MUST return the table of the Address Creation Templates available.
        /// Specifying this flag causes the NSPI server to ignore the NspiUnicodeStrings flag.
        /// </summary>
        NspiAddressCreationTemplates = 0x00000002,

        /// <summary>
        /// Specifies that the NSPI server MUST return all strings as Unicode representations
        /// rather than as multibyte strings in the client's codepage.
        /// </summary>
        NspiUnicodeStrings = 0x00000004
    }

    /// <summary>
    /// These values are used to specify optional behavior to an NSPI server. They appear as bit Flags in methods that return property values to the client (NspiGetPropList, NspiGetProps, and NspiQueryRows).
    /// </summary>
    public enum RetrievePropertyFlags : uint
    {
        /// <summary>
        /// Client requires that the server MUST NOT include proptags with the Property Type PtypEmbeddedTable in any lists of proptags that the server creates on behalf of the client.
        /// </summary>
        fSkipObjects = 0x00000001,

        /// <summary>
        /// Client requires that the server MUST return Entry ID values in Ephemeral Entry ID form.
        /// </summary>
        fEphID = 0x00000002
    }

    /// <summary>
    /// Specifies operation struct. 
    /// </summary>
    public struct Operation
    {
        /// <summary>
        /// Operation Type.
        /// </summary>
        public uint OperationInstruction;
        
        /// <summary>
        /// First operand.
        /// </summary>
        public uint? PropTag1;
        
        /// <summary>
        /// Second operand
        /// </summary>
        public uint? PropTag2;
        
        /// <summary>
        /// Jump offset.
        /// </summary>
        public uint? Offset;
    }

    /// <summary>
    /// An encoding of the PropTagArray data structure defined in [MS-OXCDATA].
    /// </summary>
    public struct PropertyTagArray_r
    {
        /// <summary>
        /// Encodes the Count field of PropTagArray. This field MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// Encodes the PropertyTags field of PropTagArray.
        /// </summary>
        [Size("Values")]
        public uint[] AulPropTag;
    }

    /// <summary>
    /// Varies depending on the control.
    /// </summary>
    public struct CNTRL
    {
        /// <summary>
        /// Varies depending on the control. 
        /// </summary>
        public uint Type;
       
        /// <summary>
        /// Varies depending on the control. 
        /// </summary>
        public uint Size;
       
        /// <summary>
        /// The offset in BYTEs from the base of the TRowSet structure to a null-terminated non-Unicode string. 
        /// </summary>
        public uint String;
       
        /// <summary>
        /// The string value of the control.
        /// </summary>
        public string StringValue;
    }

    /// <summary>
    /// Describes a control that MUST be presented to the user in a display area. 
    /// </summary>
    public struct TROW
    {
        /// <summary>
        /// X coordinate of the upper-left corner of the control.
        /// </summary>
        public uint XPos;
       
        /// <summary>
        /// Width of the control.
        /// </summary>
        public uint DeltaX;
       
        /// <summary>
        /// Y coordinate of the upper- left corner of the control.
        /// </summary>
        public uint YPos;
        
        /// <summary> 
        /// Height of the control.
        /// </summary>
        public uint DeltaY;
       
        /// <summary>
        /// Type of the control.
        /// </summary>
        public uint ControlType;
       
        /// <summary>
        /// Flags that describe the control's attributes.
        /// </summary>
        public uint ControlFlags;
       
        /// <summary>
        /// Structure that contains data that is relevant to a particular control Type.
        /// </summary>
        public CNTRL ControlStructure;
    }

    /// <summary>
    /// Specifies TRowSet structure.
    /// </summary>
    public struct TRowSet
    {
        /// <summary>
        /// Type of the template. 
        /// </summary>
        public uint Type;
       
        /// <summary>
        /// Count of TRows that are defined in this structure.
        /// </summary>
        public uint RowCount;
       
        /// <summary>
        /// TROW array.
        /// </summary>
        [Size("RowCount")]
        public TROW[] Rows;
    }

    /// <summary>
    /// Contains the Instruction information.
    /// </summary>
    public struct ScriptInstruction
    {
        /// <summary>
        /// Specifies the number of script data that follow. 
        /// </summary>
        public uint Size;
       
        /// <summary>
        /// Specifies a series of instructions and the data that accompanies them.
        /// </summary>
        [Size("Size")]
        public uint[] ScriptData;
    }

    /// <summary>
    /// Specify the state of a table and location information.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct STAT
    {
        /// <summary>
        /// This field contains a DWORD representing a sort order.
        /// </summary>
        public uint SortType;
        
        /// <summary>
        /// This field contains a minimal entry id.
        /// The client sets this field to specify the MId of the address book container that this STAT represents
        /// </summary>
        public uint ContainerID;
       
        /// <summary>
        /// This field contains a minimal entry id. 
        /// The client sets this field to specify a beginning position in the table for the start of an NSPI method.
        /// </summary>
        public uint CurrentRec;
        
        /// <summary>
        /// This field contains a long value.
        /// The client sets this field to specify an offset from the beginning position in the table for the start of an NSPI method.
        /// </summary>
        public int Delta;
        
        /// <summary>
        /// This field contains a DWORD value specifying a position in the table.
        /// </summary>
        public uint NumPos;
       
        /// <summary>
        /// This field contains a DWORD specifying the number of rows in the table.
        /// </summary>
        public uint TotalRecs;
       
        /// <summary>
        /// This field contains a DWORD value representing a codepage.
        /// </summary>
        public uint CodePage;
        
        /// <summary>
        /// This field contains a DWORD value representing a language code identifier (LCID).
        /// </summary>
        public uint TemplateLocale;
       
        /// <summary>
        /// This field contains a DWORD value representing an LCID.
        /// </summary>
        public uint SortLocale;
    }

    /// <summary>
    /// An encoding of the FlatUID data structure defined in [MS-OXCDATA].
    /// </summary>
    public struct FlatUID_r
    {
        /// <summary>
        /// Encodes the ordered bytes of the FlatUID data structure.
        /// </summary>
        [StaticSize(16, StaticSizeMode.Elements), Inline()]
        public byte[] Ab;
    }

    /// <summary>
    /// An encoding of an array of FlatUID_r data structures.
    /// </summary>
    public struct FlatUIDArray_r
    {
        /// <summary>
        /// The number of FlatUID_r structures represented in the FlatUIDArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The FlatUID_r data structures.
        /// </summary>
        [Size("Values")]
        public FlatUID_r[] Guid;
    }

    /// <summary>
    /// An encoding of an array of references to Unicode strings.
    /// </summary>
    public struct WStringArray_r
    {
        /// <summary>
        /// The number of Unicode character strings references represented in the WStringArray_r structure.
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The Unicode character string references. The strings referred to are NULL-terminated.
        /// </summary>
        [Size("Values"), String(StringEncoding.Unicode)]
        public string[] LppszW;
    }

    /// <summary>
    /// An encoding of an array of FILETIME structures.
    /// </summary>
    public struct DateTimeArray_r
    {
        /// <summary>
        /// The number of FILETIME data structures represented in the DateTimeArray_r structure.
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The FILETIME data structures.
        /// </summary>
        [Size("Values")]
        public FILETIME[] Lpft;
    }

    /// <summary>
    /// An encoding of an array of 16-bit integers.
    /// </summary>
    public struct ShortArray_r
    {
        /// <summary>
        /// The number of 16-bit integer values represented in the ShortArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The 16-bit integer values.
        /// </summary>
        [Size("Values")]
        public short[] Lpi;
    }

    /// <summary>
    /// An encoding of an array of 32-bit integers.
    /// </summary>
    public struct LongArray_r
    {
        /// <summary>
        /// The number of 32-bit integers represented in this structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The 32-bit integer values.
        /// </summary>
        [Size("Values")]
        public int[] Lpl;
    }

    /// <summary>
    /// An encoding of an array of references to 8-bit character strings.
    /// </summary>
    public struct StringArray_r
    {
        /// <summary>
        /// The number of 8-bit character strings references represented in the StringArray_r structure. 
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
       
        /// <summary>
        /// The 8-bit character string references. The strings referred to are NULL-terminated.
        /// </summary>
        [Size("Values"), String(StringEncoding.ASCII)]
        public string[] LppszA;
    }

    /// <summary>
    /// An encoding of an array of non-interpreted bytes.
    /// </summary>
    public struct Binary_r
    {
        /// <summary>
        /// The number of non-interpreted bytes represented in this structure.
        /// This value MUST NOT exceed 2,097,152.
        /// </summary>
        public uint Cb;
        
        /// <summary>
        /// The non-interpreted bytes.
        /// </summary>
        [Size("Cb")]
        public byte[] Lpb;
    }

    /// <summary>
    /// An array of Binary_r data structures.
    /// </summary>
    public struct BinaryArray_r
    {
        /// <summary>
        /// The number of Binary_r data structures represented in the BinaryArray_r structure.
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Values;
        
        /// <summary>
        /// The Binary_r data structures.
        /// </summary>
        [Size("Values")]
        public Binary_r[] Lpbin;
    }

    /// <summary>
    /// A 64-bit value that represents the number of 100-nanosecond intervals
    /// that have elapsed since January 1, 1601, Coordinated Universal Time (UTC).
    /// </summary>
    public struct FILETIME
    {
        /// <summary>
        /// Specifies the low 32 bits of the FILETIME.
        /// </summary>
        public uint LowDateTime;
        
        /// <summary>
        /// Specifies the high 32 bits of the FILETIME.
        /// </summary>
        public uint HighDateTime;
    }

    /// <summary>
    /// An encoding of a single instance of any Type of property value.
    /// </summary>
    [Union("System.int")]
    public struct PROP_VAL_UNION
    {
        /// <summary>
        /// An encoding of the value of a property that can contain a single 16-bit integer value.
        /// </summary>
        [Case("0x00000002")]
        public short I;
        
        /// <summary>
        /// An encoding of the value of a property that can contain a single 32-bit integer value.
        /// </summary>
        [Case("0x00000003")]
        public int L;
        
        /// <summary>
        /// An encoding of the value of a property that can contain a single Boolean value.
        /// </summary>
        [Case("0x0000000B")]
        public ushort B;
        
        /// <summary>
        /// An encoding of the value of a property that can contain a single 8-bit character string value.
        /// </summary>
        [Case("0x0000001E")]
        [String(StringEncoding.ASCII)]
        public string LpszA;
       
        /// <summary>
        /// An encoding of the value of a property that can contain a single binary data value.
        /// </summary>
        [Case("0x00000102")]
        public Binary_r Bin;
       
        /// <summary>
        /// An encoding of the value of a property that can contain a single Unicode string value.
        /// </summary>
        [Case("0x0000001F")]
        [String(StringEncoding.Unicode)]
        public string LpszW;
       
        /// <summary>
        /// An encoding of the value of a property that can contain a single GUID value.
        /// </summary>
        [Case("0x00000048")]
        [StaticSize(1, StaticSizeMode.Elements)]
        public FlatUID_r[] Guid;
       
        /// <summary>
        /// An encoding of the value of a property that can contain a single 64-bit integer value.
        /// </summary>
        [Case("0x00000040")]
        public FILETIME FileTime;
      
        /// <summary>
        /// An encoding of the value of a property that can contain a single PtypErrorCode value.
        /// </summary>
        [Case("0x0000000A")]
        public int ErrorCode;
       
        /// <summary>
        /// An encoding of the values of a property that can contain multiple 16-bit integer values.
        /// </summary>
        [Case("0x00001002")]
        public ShortArray_r MVi;
       
        /// <summary>
        /// An encoding of the values of a property that can contain multiple 32-bit integer values.
        /// </summary>
        [Case("0x00001003")]
        public LongArray_r MVl;
       
        /// <summary>
        /// An encoding of the values of a property that can contain multiple 8-bit character string values.
        /// </summary>
        [Case("0x0000101E")]
        public StringArray_r MVszA;
     
        /// <summary>
        /// An encoding of the values of a property that can contain multiple binary data values.
        /// </summary>
        [Case("0x00001102")]
        public BinaryArray_r MVbin;
      
        /// <summary>
        /// An encoding of the values of a property that can contain multiple GUID values.
        /// </summary>
        [Case("0x00001048")]
        public FlatUIDArray_r MVguid;
      
        /// <summary>
        /// An encoding of the values of a property that can contain multiple Unicode string values.
        /// </summary>
        [Case("0x0000101F")]
        public WStringArray_r MVszW;
       
        /// <summary>
        /// An encoding of the value of a property that can contain multiple 64-bit integer values.
        /// </summary>
        [Case("0x00001040")]
        public DateTimeArray_r MVft;
      
        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        [Case("0x00000001, 0x0000000D")]
        public int Reserved;
    }

    /// <summary>
    /// An encoding of the PropertyValue data structure defined in [MS-OXCDATA].
    /// </summary>
    public struct PropertyValue_r
    {
        /// <summary>
        /// Encodes the property tag of the property whose value is represented by the PropertyValue_r data structure.
        /// </summary>
        public uint PropTag;
       
        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;
      
        /// <summary>
        /// Encodes the actual value of the property represented by the PropertyValue_r data structure.
        /// </summary>
        [Switch("ulPropTag & 0x0000FFFF")]
        public PROP_VAL_UNION Value;
    }

    /// <summary>
    /// An encoding of the StandardPropertyRow data structure defined in [MS-OXCDATA].
    /// </summary>
    public struct PropertyRow_r
    {
        /// <summary>
        /// Reserved. All clients and servers MUST set this value to the constant 0x00000000.
        /// </summary>
        public uint Reserved;
       
        /// <summary>
        /// The number of PropertyValue_r structures represented in the PropertyRow_r structure.
        /// </summary>
        public uint Values;
       
        /// <summary>
        /// Encodes the ValueArray field of the StandardPropertyRow data structure.
        /// </summary>
        [Size("cValues")]
        public PropertyValue_r[] Props;
    }

    /// <summary>
    /// An encoding of the StandardPropertyRow data structure defined in [MS-OXCDATA].
    /// </summary>
    public struct PropertyRowSet_r
    {
        /// <summary>
        /// Encodes the RowCount field of the PropertyRowSet data structures.
        /// This value MUST NOT exceed 100,000.
        /// </summary>
        public uint Rows;
       
        /// <summary>
        /// Encodes the Rows field of the PropertyRowSet data structure.
        /// </summary>
        [Size("cRows"), Inline()]
        public PropertyRow_r[] PropertyRowSet;
    }
}