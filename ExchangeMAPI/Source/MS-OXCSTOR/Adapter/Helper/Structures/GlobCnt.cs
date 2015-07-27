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
    /// Global Change
    /// </summary>
    public class GlobCnt
    {
        /// <summary>
        /// Command type
        /// </summary>
        private CommandType type;

        /// <summary>
        /// Base command
        /// </summary>
        private BaseCommand command;

        /// <summary>
        /// Gets or sets the type
        /// </summary>
        public CommandType Type
        {
            get
            {
                return this.type;
            }

            set
            {
                this.type = value;
            }
        }

        /// <summary>
        /// Gets or sets the command
        /// </summary>
        public BaseCommand Command
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
    }
}