namespace Microsoft.Protocols.TestSuites.MS_ASWBXML
{
    /// <summary>
    /// Global WBXML tokens
    /// </summary>
    internal enum GlobalTokens
    {
        /// <summary>
        /// Change the code page for the current token state. Followed by a single u_int8 indicating the new code page number.
        /// </summary>
        SWITCH_PAGE = 0x00,

        /// <summary>
        /// Indicates the end of an attribute list or the end of an element.
        /// </summary>
        END = 0x01,

        /// <summary>
        /// A character entity. Followed by an mb_u_int32 encoding the character entity number.
        /// </summary>
        ENTITY = 0x02,

        /// <summary>
        /// Inline string. Followed by a charset-dependent string with termination.
        /// </summary>
        STR_I = 0x03,

        /// <summary>
        /// An unknown tag or attribute name. Followed by an mb_u_int32 that encodes an offset into the string table.
        /// </summary>
        LITERAL = 0x04,

        /// <summary>
        /// Inline string document-type-specific extension token. Token is followed by a charset-dependent string with termination.
        /// </summary>
        EXT_I_0 = 0x40,

        /// <summary>
        /// Inline string document-type-specific extension token. Token is followed by a charset-dependent string with termination.
        /// </summary>
        EXT_I_1 = 0x41,

        /// <summary>
        /// Inline string document-type-specific extension token. Token is followed by a charset-dependent string with termination.
        /// </summary>
        EXT_I_2 = 0x42,

        /// <summary>
        /// Processing instruction.
        /// </summary>
        PI = 0x43,

        /// <summary>
        /// Unknown tag, with content.
        /// </summary>
        LITERAL_C = 0x44,

        /// <summary>
        /// Inline integer document-type-specific extension token. Token is followed by an mb_u_int32.
        /// </summary>
        EXT_T_0 = 0x80,

        /// <summary>
        /// Inline integer document-type-specific extension token. Token is followed by an mb_u_int32.
        /// </summary>
        EXT_T_1 = 0x81,

        /// <summary>
        /// Inline integer document-type-specific extension token. Token is followed by an mb_u_int32.
        /// </summary>
        EXT_T_2 = 0x82,

        /// <summary>
        /// String table reference. Followed by an mb_u_int32 encoding a byte offset from the beginning of the string table.
        /// </summary>
        STR_T = 0x83,

        /// <summary>
        /// Unknown tag, with attributes.
        /// </summary>
        LITERAL_A = 0x84,

        /// <summary>
        /// Single-byte document-type-specific extension token.
        /// </summary>
        EXT_0 = 0xC0,

        /// <summary>
        /// Single-byte document-type-specific extension token.
        /// </summary>
        EXT_1 = 0xC1,

        /// <summary>
        /// Single-byte document-type-specific extension token.
        /// </summary>
        EXT_2 = 0xC2,

        /// <summary>
        /// Opaque document-type-specific data.
        /// </summary>
        OPAQUE = 0xC3,

        /// <summary>
        /// Unknown tag, with content and attributes.
        /// </summary>
        LITERAL_AC = 0xC4
    }
}