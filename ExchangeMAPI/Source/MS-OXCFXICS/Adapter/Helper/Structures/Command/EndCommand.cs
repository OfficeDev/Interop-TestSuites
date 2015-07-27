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
    /// Represent an end command.
    /// </summary>
    public class EndCommand : Command
    {
        /// <summary>
        /// Initializes a new instance of the EndCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        public EndCommand(byte command)
            : base(command, 0x00, 0x00)
        { 
        }
    }
}