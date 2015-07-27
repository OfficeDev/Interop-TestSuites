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
    /// Two elements tuple.
    /// </summary>
    /// <typeparam name="T1">The type of the 1st element.</typeparam>
    /// <typeparam name="T2">The type of the 2nd element.</typeparam>
    public class Tuple<T1, T2>
    {
        /// <summary>
        /// The first item.
        /// </summary>
        private T1 item1;

        /// <summary>
        /// The second item.
        /// </summary>
        private T2 item2;

        /// <summary>
        /// Initializes a new instance of the Tuple class.
        /// </summary>
        /// <param name="item1">The 1st element.</param>
        /// <param name="item2">The 2nd element.</param>
        public Tuple(T1 item1, T2 item2)
        {
            this.item1 = item1;
            this.item2 = item2;
        }

        /// <summary>
        /// Gets or sets the 1st element.
        /// </summary>
        public T1 Item1
        {
            get { return this.item1; }
            set { this.item1 = value; }
        }

        /// <summary>
        /// Gets or sets the 2nd element.
        /// </summary>
        public T2 Item2
        {
            get { return this.item2; }
            set { this.item2 = value; }
        }
    }
}