//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    /// <summary>
    /// Define the Helper class.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// Prevents a default instance of the AdapterHelper class from being created.
        /// </summary>
        private AdapterHelper()
        {
        }

        /// <summary>
        /// Check the character in specific string whether is valid ASCII character.
        /// </summary>
        /// <param name="strMsg">A string value that is checked.</param>
        /// <param name="minChar">The minimum valid ASCII character.</param>
        /// <param name="maxChar">The maximum valid ASCII character.</param>
        /// <returns>If all characters in specific string are valid, return true, else return false.</returns>
        public static bool IsStringValueValid(string strMsg, char minChar, char maxChar)
        {
            bool isValid = true;
            for (int index = 0; index < strMsg.Length; index++)
            {
                if (strMsg[index] < minChar || strMsg[index] > maxChar)
                {
                    isValid = false;
                    break;
                }
            }

            return isValid;
        }

        /// <summary>
        /// Check the length of specific string whether is valid value.
        /// </summary>
        /// <param name="strMsg">A string value that is checked.</param>
        /// <param name="minLength">An integer value indicates a length that the string valid length must greater than.</param>
        /// <param name="maxLength">An integer value indicates a length that the string valid length must smaller than.</param>
        /// <returns>If the length specific string is valid, return true, else return false.</returns>
        public static bool IsStringLengthValid(string strMsg, int minLength, int maxLength)
        {
            bool isValid = strMsg.Length > minLength && strMsg.Length < maxLength;
            return isValid;
        }
    }
}