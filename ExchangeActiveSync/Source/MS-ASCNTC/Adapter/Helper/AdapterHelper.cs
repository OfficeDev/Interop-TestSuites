namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System;
    using System.Drawing;
    using System.IO;

    /// <summary>
    /// The class provides the methods to assist MS-ASCNTCAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary> 
        /// Check whether the base64 encoded string contains a picture.
        /// </summary> 
        /// <param name="base64Value">A base64 encoded string which should contain the picture data.</param> 
        /// <returns>True, if the base64 encoded string contains a picture. Otherwise, False.</returns> 
        public static bool IsPicture(string base64Value)
        {
            try
            {
                byte[] bytes = Convert.FromBase64String(base64Value);
                return IsPicture(bytes);
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        /// <summary>
        /// Check whether the byte array contains a picture.
        /// </summary>
        /// <param name="value">The byte array which should contain the picture data.</param>
        /// <returns>True, if the byte array contains a picture. Otherwise, False.</returns>
        public static bool IsPicture(byte[] value)
        {
            MemoryStream memStream = new MemoryStream(value);
            try
            {
                Image.FromStream(memStream);

                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
            finally
            {
                memStream.Dispose();
            }
        }
    }
}