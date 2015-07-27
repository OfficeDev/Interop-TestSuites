//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    /// <summary>
    /// Organizes CONST values which are used in adapter.
    /// </summary>
    public static class ConstValues
    {
        /// <summary>
        /// According to section 2.2.9.2.1 of MS-OXCROPS, MaximumByteCount field is present when ByteCount is equal to 0xBABE(47806).
        /// </summary>
        public const ushort StreamLengthByteCountIndicator = 0xBABE;

        /// <summary>
        /// Indicates the Associated Flag is none.
        /// </summary>
        public const byte RopCreateMessageAssociatedFlagNone = 0x00;

        /// <summary>
        /// Used in filling property bytes.
        /// </summary>
        public const byte PropertyBytesFilling = 0xF;

        /// <summary>
        /// Indicates OpenModeFlags is ReadOnly.
        /// </summary>
        public const byte OpenModeFlagsReadOnly = 0x00;

        /// <summary>
        /// Indicates OpenModeFlags is ReadWrite.
        /// </summary>
        public const byte OpenModeFlagsReadWrite = 0x01;

        /// <summary>
        /// Indicates OpenModeFlags is Create.
        /// </summary>
        public const byte OpenModeFlagsCreate = 0x02;

        /// <summary>
        /// Indicates OpenModeFlags is BestAccess.
        /// </summary>
        public const byte OpenModeFlagsBestAccess = 0x03;

        /// <summary>
        /// Indicates the offset is zero.
        /// </summary>
        public const long RopSeekStreamOffsetZero = 0;

        /// <summary>
        /// No limit for property size.
        /// </summary>
        public const ushort PropertySizeLimitNone = 0;

        /// <summary>
        /// Indicates no use UNICODE.
        /// </summary>
        public const ushort WantUnicodeNo = 0;

        /// <summary>
        /// Key which is defined for a folder.
        /// </summary>
        public const string ServerObjectTypeFolderKey = "Folder1";

        /// <summary>
        /// Key which is defined for a message.
        /// </summary>
        public const string ServerObjectTypeMessageKey = "Message1";

        /// <summary>
        /// Key which is defined for an attachment.
        /// </summary>
        public const string ServerObjectTypeAttachmentKey = "Attachment1";

        /// <summary>
        /// Size of PtypTime is 8 in byte.
        /// </summary>
        public const int PtypTimeSizeInByte = 8;

        /// <summary>
        /// The property count for test the order of response of RopGetPropertyIdsFromNamesMethod.
        /// </summary>
        public const int TestPropertyCount = 3;
    }
}