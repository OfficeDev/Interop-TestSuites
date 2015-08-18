namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Property Data Type Enumeration
    /// </summary>
    //// [CLSCompliant(false)]
    public enum PropertyDataType : ushort
    {      
        /// <summary>
        ///  2 bytes, a 16-bit integer 
        /// </summary>
        PtypInteger16 = 0x0002,

        /// <summary>
        ///  4 bytes, a 32-bit integer 
        /// </summary>
        PtypInteger32 = 0x0003,

        /// <summary>
        ///  4 bytes, a 32-bit floating point number 
        /// </summary>
        PtypFloating32 = 0x0004,

        /// <summary>
        ///  8 bytes, a 64-bit floating point number 
        /// </summary>
        PtypFloating64 = 0x0005,

        /// <summary>
        ///  8 bytes, a 64-bit signed, scaled integer representation of a decimal currency value, with 
        /// 4 places to the right of the decimal point 
        /// </summary>
        PtypCurrency = 0x0006,

        /// <summary>
        ///  8 bytes, a 64-bit floating point number in which the whole number part represents the number 
        /// of days since December 30, 1899, and the fractional part represents the fraction of a day since 
        /// midnight 
        /// </summary>
        PtypFloatingTime = 0x0007,

        /// <summary>
        ///  4 bytes, a 32-bit integer encoding error information as specified in section 2.4.1. 
        /// </summary>
        PtypErrorCode = 0x000A,

        /// <summary>
        ///  1 byte, restricted to 1 or 0 
        /// </summary>
        PtypBoolean = 0x000B,

        /// <summary>
        ///  8 bytes, a 64-bit integer 
        /// </summary>
        PtypInteger64 = 0x0014,

        /// <summary>
        ///  Variable size, a string of Unicode characters in UTF-16LE encoding with terminating null character 
        /// (0x0000). 
        /// </summary>
        PtypString = 0x001F,

        /// <summary>
        ///  Variable size, a string of multi-byte characters in externally specified encoding with terminating 
        /// null character (single 0 byte). 
        /// </summary>
        PtypString8 = 0x001E,

        /// <summary>
        ///  8 bytes, a 64-bit integer representing the number of 100-nanosecond intervals since January 
        /// 1, 1601 
        /// </summary>
        PtypTime = 0x0040,

        /// <summary>
        ///  16 bytes, a GUID with Data1, Data2, and Data3 fields in little-endian format 
        /// </summary>
        PtypGuid = 0x0048,

        /// <summary>
        ///  Variable size, a 16-bit COUNT followed by a structure as specified in section 2.11.1.3. 
        /// </summary>
        PtypServerId = 0x00FB,

        /// <summary>
        ///  Variable size, a byte array representing one or more Restriction structures as specified in 
        /// section 2.12. 
        /// </summary>
        PtypRestriction = 0x00FD,

        /// <summary>
        ///  Variable size, a 16-bit COUNT of actions (not bytes) followed by that many Rule Action structures, 
        /// as specified in [MS-OXORULE] section 2.2.5. 
        /// </summary>
        PtypRuleAction = 0x00FE,

        /// <summary>
        ///  Variable size, a COUNT followed by that many bytes. 
        /// </summary>
        PtypBinary = 0x0102,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypInteger16 values. 
        /// </summary>
        PtypMultipleInteger16 = 0x1002,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypInteger32 values. 
        /// </summary>
        PtypMultipleInteger32 = 0x1003,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypFloating32 values. 
        /// </summary>
        PtypMultipleFloating32 = 0x1004,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypFloating64 values. 
        /// </summary>
        PtypMultipleFloating64 = 0x1005,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypCurrency values. 
        /// </summary>
        PtypMultipleCurrency = 0x1006,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypFloatingTime values. 
        /// </summary>
        PtypMultipleFloatingTime = 0x1007,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypInteger64 values. 
        /// </summary>
        PtypMultipleInteger64 = 0x1014,

        /// <summary>
        ///  Variable size, a COUNT followed by that PtypString values. 
        /// </summary>
        PtypMultipleString = 0x101F,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypString8 values. 
        /// </summary>
        PtypMultipleString8 = 0x101E,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypTime values. 
        /// </summary>
        PtypMultipleTime = 0x1040,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypGuid values. 
        /// </summary>
        PtypMultipleGuid = 0x1048,

        /// <summary>
        ///  Variable size, a COUNT followed by that many PtypBinary values. 
        /// </summary>
        PtypMultipleBinary = 0x1102,

        /// <summary>
        ///  Any: this property type value matches any type; a server MUST return the actual type in its 
        /// response. Servers MUST NOT return this type in response to a client request other than NspiGetIDsFromNames 
        /// or RopGetPropertyIdsFromNames ([MS-OXCROPS]). 
        /// </summary>
        PtypUnspecified = 0x0000,

        /// <summary>
        ///  None: This property is a placeholder. 
        /// </summary>
        PtypNull = 0x0001,

        /// <summary>
        ///  The property value is a COM object, as specified in section 2.11.1.4, or PtypEmbeddedTable.
        /// </summary>
        PtypObject = 0x000D,

        /// <summary>
        /// Special type ID for unsigned
        /// </summary>
        unsigned = 0xFFFF,
    }
}