//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    /// <summary>
    /// This is an 8-bit structure used as both TableFlags and ModifyFlags structure.
    /// </summary>
    public class RequestBufferFlags
    {
        /// <summary>
        /// isReservedBitsSet property
        /// </summary>
        private bool isReservedBitsSet = false;

        /// <summary>
        /// isReplaceRowsFlagSet Property
        /// </summary>
        private bool isReplaceRowsFlagSet = false;

        /// <summary>
        /// isIncludeFreeBusyFlagSet property
        /// </summary>
        private bool isIncludeFreeBusyFlagSet = false;

        /// <summary>
        /// bufferFlags property
        /// </summary>
        private byte bufferFlags = 0x00;

        /// <summary>
        /// Gets or sets a value indicating whether to set the ReservedBits or not, 
        /// if true, all reserved bits will be set as 1;
        /// if false, all reserved bits will be set as 0.
        /// </summary>
        public bool IsReservedBitsSet
        {
            get
            {
                return this.isReservedBitsSet;
            }

            set
            {
                this.isReservedBitsSet = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to set the ReplaceRows flag or not, 
        /// if true, ReplaceRows flag will be set as 1;
        /// if false, ReplaceRows flag will be set as 0.
        /// </summary>
        public bool IsReplaceRowsFlagSet
        {
            get { return this.isReplaceRowsFlagSet; }
            set { this.isReplaceRowsFlagSet = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to set the IncludeFreeBusy flag or not,
        /// if true, IncludeFreeBusy flag will be set as 1;
        /// if false, IncludeFreeBusy flag will be set as 0;
        /// </summary>
        public bool IsIncludeFreeBusyFlagSet
        {
            get { return this.isIncludeFreeBusyFlagSet; }
            set { this.isIncludeFreeBusyFlagSet = value; }
        }

        /// <summary>
        /// Gets 8-bits structure property
        /// </summary>
        public byte BufferFlags
        {
            get 
            {
                this.bufferFlags = 0x00;
                if (this.isIncludeFreeBusyFlagSet)
                {
                    this.bufferFlags |= 0x02;
                }

                if (this.isReplaceRowsFlagSet)
                {
                    this.bufferFlags |= 0x01;
                }

                if (this.isReservedBitsSet)
                {
                    this.bufferFlags |= 0xFC;
                }

                return this.bufferFlags;
            }  
        }
    }
}