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
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents a command in GLOBSET serialization.
    /// </summary>
    public abstract class Command
    {
        /// <summary>
        /// A byte indicates command type.
        /// </summary>
        private byte commandByte;

        /// <summary>
        /// A GLOBCNT List deserialized by this command.
        /// </summary>
        private List<GLOBCNT> correspondingGLOBCNTList;

        /// <summary>
        /// A list contains  GLOBCNTRanges which have been deserialized from a stream.
        /// </summary>
        private List<GLOBCNTRange> correspondingGLOBCNTRangeList;

        /// <summary>
        /// Initializes a new instance of the Command class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        /// <param name="low">Low value of the command byte interval.</param>
        /// <param name="high">High value of the command byte interval.</param>
        protected Command(byte command, byte low, byte high)
        {
            this.correspondingGLOBCNTList = null;
            if (!this.CheckCommand(command, low, high))
            {
                AdapterHelper.Site.Assert.Fail("The command is invalid.");
            }

            this.commandByte = command;
        }

        /// <summary>
        /// Gets or sets the GLOBCNTRange List deserialized by this command
        /// </summary>
        public List<GLOBCNTRange> CorrespondingGLOBCNTRangeList
        {
            get { return this.correspondingGLOBCNTRangeList; }
            set { this.correspondingGLOBCNTRangeList = value; }
        }

        /// <summary>
        /// Gets a GLOBCNT List deserialized by this command.
        /// </summary>
        public List<GLOBCNT> CorrespondingGLOBCNTList
        {
            get 
            {
                this.correspondingGLOBCNTList = new List<GLOBCNT>();
                if (this.correspondingGLOBCNTRangeList != null)
                {
                    this.correspondingGLOBCNTList = GLOBSET.GetGLOBCNTList(this.correspondingGLOBCNTRangeList);
                }

                return this.correspondingGLOBCNTList;
            }
        }

        /// <summary>
        /// Gets a byte indicate the type of this command.
        /// </summary>
        public byte CommandByte
        {
            get { return this.commandByte; }
        }

        /// <summary>
        /// Indicate whether the command byte is in an interval.
        /// </summary>
        /// <param name="command">The command byte.</param>
        /// <param name="low">Low value of the interval.</param>
        /// <param name="high">High value of the interval.</param>
        /// <returns>If the command byte in an interval, return true, else false.</returns>
        protected bool CheckCommand(byte command, byte low, byte high)
        {
            return command >= low && command <= high;
        }
    }
}