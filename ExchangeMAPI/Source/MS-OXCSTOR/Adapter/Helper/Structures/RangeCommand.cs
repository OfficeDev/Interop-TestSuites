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
    using System;

    /// <summary>
    /// RangeCommand class
    /// </summary>
    public class RangeCommand : BaseCommand
    {
        /// <summary>
        /// Private variable CommandBytesLength
        /// </summary>
        private const int CommandBytesLength = 2;

        /// <summary>
        /// Initializes a new instance of the RangeCommand class
        /// </summary>
        public RangeCommand()
        {
            this.Command = (byte)CommandType.RangCommand;
        }

        /// <summary>
        /// Generate random bytes in command
        /// </summary>
        public void GenerateRangeBytes()
        {
            this.CommandBytes = new byte[CommandBytesLength];
            this.CommandBytes[0] = 0xc1;
            this.CommandBytes[1] = 0xc2;
        }

        /// <summary>
        /// Get the size of the RangeCommand
        /// </summary>
        /// <returns>The size of the RangeCommand</returns>
        public override int Size()
        {
            // 1 byte is for command, 1 byte is for low value, and 1 byte is for high value
            return 3;
        }

        /// <summary>
        /// Get the bytes of the RangeCommand
        /// </summary>
        /// <returns>The bytes of the RangeCommand</returns>
        public override byte[] GetBytes()
        {
            byte[] resultBytes = new byte[1 + CommandBytesLength];
            resultBytes[0] = this.Command;
            Array.Copy(this.CommandBytes, 0, resultBytes, 1, this.CommandBytes.Length);
            return resultBytes;
        }
    }
}