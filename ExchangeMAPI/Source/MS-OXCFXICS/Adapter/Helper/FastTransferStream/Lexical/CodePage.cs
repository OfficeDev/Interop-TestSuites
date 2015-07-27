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

    /// <summary>
    /// The CodePage class used to transmit string properties using the code page format.
    /// </summary>
    public class CodePage
    {
        /// <summary>
        /// Store the information which are combine with the property A and the property CodePageId.
        /// </summary>
        private ushort codePageData;

        /// <summary>
        /// Initializes a new instance of the CodePage class.
        /// </summary>
        public CodePage()
        {
            this.A = 0x8000;
        }

        /// <summary>
        /// Gets or sets the value specifies the property is an internal code page string.
        /// </summary>
        public ushort A
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the value specifies code page ID identifier for the code page used to encode the string property.
        /// </summary>
        public ushort CodePageId
        {
            get;
            set;
        }

        /// <summary>
        /// Get the size of the code page structure.
        /// </summary>
        /// <returns>Returns the size of code page structure.</returns>
        public int Size()
        {
            return sizeof(ushort);
        }

        /// <summary>
        /// Deserialize the code page structure.
        /// </summary>
        /// <param name="data">The code page data stores in the ushort type parameter.</param>
        public void Deserialize(ushort data)
        {
            this.codePageData = data;
            if (this.VerifyPropertyA(this.codePageData))
            {
                this.A = 0x8000;
                this.CodePageId = (ushort)(this.codePageData - this.A);
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The field A of Code Page property is not set to 1.");
            }
        }

        /// <summary>
        /// Serialize the code page data.
        /// </summary>
        /// <returns>Format the code page structure to byte array.</returns>
        public byte[] Serialize()
        {
            if (this.VerifyPropertyA(this.A) && this.CodePageId < 0x8000)
            {
                this.codePageData = (ushort)(this.A + this.CodePageId);
                byte[] data = BitConverter.GetBytes(this.codePageData);
                return data;
            }
            else
            {
                AdapterHelper.Site.Assert.Fail("The top bit of the property A is not set to 1, or the property CodePageId is out of scope.");
                return null;
            }
        }

        /// <summary>
        /// Verify the value of property A whether sets the top bit for 1.
        /// </summary>
        /// <param name="propertyA">The value of property A.</param>
        /// <returns>Returns the verified result.</returns>
        private bool VerifyPropertyA(ushort propertyA)
        {
            ushort propertyValue = propertyA;
            int result = propertyValue & 0x8000;
            if (result == 0x00008000)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}