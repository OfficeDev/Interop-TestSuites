//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASWBXML
{
    using System.Collections.Generic;

    /// <summary>
    /// Represents a code page.
    /// </summary>
    internal class CodePage
    {
        /// <summary>
        /// A dictionary stores token-tag pair
        /// </summary>
        private Dictionary<byte, string> tokenTag = new Dictionary<byte, string>();

        /// <summary>
        /// Gets or sets the namespace of the token in this code page
        /// </summary>
        public string Namespace
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the xmlns prefix of the token in this code page
        /// </summary>
        public string Xmlns
        {
            get;
            set;
        }

        /// <summary>
        /// Adds token tag pair.
        /// </summary>
        /// <param name="token">The token to add</param>
        /// <param name="tag">The tag to add</param>
        public void AddToken(byte token, string tag)
        {
            this.tokenTag.Add(token, tag);
        }

        /// <summary>
        /// Gets the token by tag name.
        /// </summary>
        /// <param name="tag">The tag name of the token</param>
        /// <returns>The token corresponding to the tag</returns>
        public byte GetToken(string tag)
        {
            if (this.tokenTag.ContainsValue(tag))
            {
                foreach (byte token in this.tokenTag.Keys)
                {
                    if (this.tokenTag[token].Equals(tag, System.StringComparison.CurrentCultureIgnoreCase))
                    {
                        return token;
                    }
                }
            }

            return 0xFF;
        }

        /// <summary>
        /// Gets the tag name by token
        /// </summary>
        /// <param name="token">The token of the tag</param>
        /// <returns>The tag name corresponding to the token</returns>
        public string GetTag(byte token)
        {
            if (this.tokenTag.ContainsKey(token))
            {
                return this.tokenTag[token];
            }

            return null;
        }
    }
}