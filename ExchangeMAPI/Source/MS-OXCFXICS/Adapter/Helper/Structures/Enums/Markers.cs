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
    /// <summary>
    /// Syntactical markers
    /// </summary>
    public enum Markers : uint
    {
        // Folders 

        /// <summary>
        /// The PidTagStartTopFld
        /// </summary>
        PidTagStartTopFld = 0x40090003, 

        /// <summary>
        /// The PidTagEndFolder
        /// </summary>
        PidTagEndFolder = 0x400B0003,

        /// <summary>
        /// The PidTagStartSubFld
        /// </summary>
        PidTagStartSubFld = 0x400A0003,
        
        // Messages and their parts

        /// <summary>
        /// The PidTagStartMessage
        /// </summary>
        PidTagStartMessage = 0x400C0003, 

        /// <summary>
        /// The PidTagEndMessage
        /// </summary>
        PidTagEndMessage = 0x400D0003,

        /// <summary>
        /// The PidTagStartFAIMsg
        /// </summary>
        PidTagStartFAIMsg = 0x40100003,

        /// <summary>
        /// The PidTagStartEmbed
        /// </summary>
        PidTagStartEmbed = 0x40010003, 

        /// <summary>
        /// The PidTagEndEmbed
        /// </summary>
        PidTagEndEmbed = 0x40020003,

        /// <summary>
        /// The PidTagStartRecip
        /// </summary>
        PidTagStartRecip = 0x40030003, 

        /// <summary>
        /// The PidTagEndToRecip
        /// </summary>
        PidTagEndToRecip = 0x40040003,

        /// <summary>
        /// The PidTagNewAttach
        /// </summary>
        PidTagNewAttach = 0x40000003, 

        /// <summary>
        /// The PidTagEndAttach
        /// </summary>
        PidTagEndAttach = 0x400E0003,

        // Synchronization download

        /// <summary>
        /// The PidTagIncrSyncChg
        /// </summary>
        PidTagIncrSyncChg = 0x40120003,

        /// <summary>
        /// The PidTagIncrSyncChgPartial
        /// </summary>
        PidTagIncrSyncChgPartial = 0x407D0003,

        /// <summary>
        /// The PidTagIncrSyncDel
        /// </summary>
        PidTagIncrSyncDel = 0x40130003,

        /// <summary>
        /// The PidTagIncrSyncEnd
        /// </summary>
        PidTagIncrSyncEnd = 0x40140003,

        /// <summary>
        /// The PidTagIncrSyncRead
        /// </summary>
        PidTagIncrSyncRead = 0x402F0003,

        /// <summary>
        /// The PidTagIncrSyncStateBegin
        /// </summary>
        PidTagIncrSyncStateBegin = 0x403A0003,

        /// <summary>
        /// The PidTagIncrSyncStateEnd
        /// </summary>
        PidTagIncrSyncStateEnd = 0x403B0003,

        /// <summary>
        /// The PidTagIncrSyncProgressMode
        /// </summary>
        PidTagIncrSyncProgressMode = 0x4074000B,

        /// <summary>
        /// The PidTagIncrSyncProgressPerMsg
        /// </summary>
        PidTagIncrSyncProgressPerMsg = 0x4075000B,

        /// <summary>
        /// The PidTagIncrSyncMessage
        /// </summary>
        PidTagIncrSyncMessage = 0x40150003,

        /// <summary>
        /// The PidTagIncrSyncGroupInfo
        /// </summary>
        PidTagIncrSyncGroupInfo = 0x407B0102,

        // Special

        /// <summary>
        /// The PidTagFXErrorInfo
        /// </summary>
        PidTagFXErrorInfo = 0x40180003,
    }

    /// <summary>
    /// Meta properties
    /// </summary>
    public enum MetaProperties
    {
        /// <summary>
        /// The PidTagEcWarning
        /// </summary>
        PidTagEcWarning = 0x400f0003,

        /// <summary>
        /// The PidTagNewFXFolder
        /// </summary>
        PidTagNewFXFolder = 0x40110102,

        /// <summary>
        /// The PidTagFXDelProp
        /// </summary>
        PidTagFXDelProp = 0x40160003,

        /// <summary>
        /// The MetaTagIncrSyncGroupId
        /// </summary>
        MetaTagIncrSyncGroupId = 0x407c0003,
        
        /// <summary>
        /// The MetaTagIncrementalSyncMessagePartial
        /// </summary>
        MetaTagIncrementalSyncMessagePartial = 0x407a0003
    }
}