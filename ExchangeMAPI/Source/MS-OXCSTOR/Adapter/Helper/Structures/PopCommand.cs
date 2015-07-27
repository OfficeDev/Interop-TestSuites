//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    /// <summary>
    /// PopCommand class
    /// </summary>
    public class PopCommand : BaseCommand
    {
        /// <summary>
        /// Initializes a new instance of the PopCommand class
        /// </summary>
        public PopCommand()
        {
            this.Command = (byte)CommandType.PopCommand;
        }

        /// <summary>
        /// Constructor size
        /// </summary>
        /// <returns>The size of the Constructor</returns>
        public override int Size()
        {
            return 1;
        }

        /// <summary>
        /// Get the bytes of the PopCommand
        /// </summary>
        /// <returns>The bytes of the PopCommand</returns>
        public override byte[] GetBytes()
        {
            byte[] resultBytes = new byte[1];
            resultBytes[0] = (byte)CommandType.RangCommand;
            return resultBytes;
        }
    }
}