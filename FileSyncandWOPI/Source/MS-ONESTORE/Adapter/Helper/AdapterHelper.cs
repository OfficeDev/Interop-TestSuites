namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;

    /// <summary>
    /// The class provides the methods to assist MS-ONESTOREAdapter.
    /// </summary>
    public class AdapterHelper
    {
        #region Adapter Help Methods
        /// <summary>
        /// This method is used to read the Guid for byte array.
        /// </summary>
        /// <param name="byteArray">The byte array.</param>
        /// <param name="startIndex">The offset of the Guid value.</param>
        /// <returns>Return the value of Guid.</returns>
        public static Guid ReadGuid(byte[] byteArray, int startIndex)
        {
            byte[] guidBuffer = new byte[16];
            Array.Copy(byteArray, startIndex, guidBuffer, 0, 16);

            return new Guid(guidBuffer);
        }

        public static uint ComputeCRC(string value)
        {
            uint crcValue = 0;

            return crcValue;
        }
        #endregion Methods
    }
}
