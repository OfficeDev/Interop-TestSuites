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
    /// The base command is used to signal the end of the GLOBSET encoding.
    /// </summary>
    public class BaseCommand
    {
        /// <summary>
        /// The command byte
        /// </summary>
        private byte command;

        /// <summary>
        /// The data in bytes contained in command
        /// </summary>
        private byte[] commandBytes;

        /// <summary>
        /// Gets or sets the command
        /// </summary>
        public byte Command
        {
            get
            {
                return this.command;
            }

            set
            {
                this.command = value;
            }
        }

        /// <summary>
        /// Gets or sets the commandBytes
        /// </summary>
        public byte[] CommandBytes
        {
            get
            {
                return this.commandBytes;
            }

            set
            {
                this.commandBytes = value;
            }
        }

        /// <summary>
        /// Get the size of command
        /// </summary>
        /// <returns>The size of command</returns>
        public virtual int Size()
        {
            return 0;
        }

        /// <summary>
        /// Get the bytes of the command
        /// </summary>
        /// <returns>The bytes of the command</returns>
        public virtual byte[] GetBytes()
        {
            return null;
        }
    }
}