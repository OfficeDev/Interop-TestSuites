//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Net.Mail;

    /// <summary>
    /// This class provides common functions to the Adapter project.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// Prevents a default instance of the AdapterHelper class from being created
        /// </summary>
        private AdapterHelper()
        { 
        }

        /// <summary>
        /// Check the format of the returned email address.
        /// </summary>
        /// <param name="returnedEmailAddr">A string contains the returned email address.</param>
        /// <returns>A Boolean value that indicates the email address format is valid or not.</returns>
        public static bool IsValidEmailAddr(string returnedEmailAddr)
        {
            try
            {
                new MailAddress(returnedEmailAddr);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        /// <summary>
        /// Get specified length byte array from specified byte array with the start index.
        /// </summary>
        /// <param name="array">The byte array which contains all data.</param>
        /// <param name="startIndex">The index where to start.</param>
        /// <param name="length">The length of expect byte array.</param>
        /// <returns>The expect byte array.</returns>
        public static byte[] GetBytes(byte[] array, int startIndex, int length)
        {
            byte[] temp = new byte[length];

            for (int i = 0; i < length; i++, startIndex++)
            {
                temp[i] = array[startIndex];
            }

            return temp;
        }

        /// <summary>
        /// Determines whether two specified byte array have the same content.
        /// </summary>
        /// <param name="array1">The first byte array.</param>
        /// <param name="array2">The second byte array.</param>
        /// <returns>true if the content of array1 is the same as the content of array2; otherwise, false.</returns>
        public static bool ByteArrayEquals(byte[] array1, byte[] array2)
        {
            if (array1 == null || array2 == null)
            {
                return true;
            }
            else if (array2.Length != array1.Length)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < array1.Length; i++)
                {
                    if (array1[i] != array2[i])
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}