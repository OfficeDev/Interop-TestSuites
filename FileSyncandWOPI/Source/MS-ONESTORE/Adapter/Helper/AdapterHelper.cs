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
        /// <summary>
        /// XOR two ExtendedGUID instances.
        /// </summary>
        /// <param name="exGuid1">The first ExtendedGUID instance.</param>
        /// <param name="exGuid2">The second ExtendedGUID instance.</param>
        /// <returns>Returns the result of XOR two ExtendedGUID instances.</returns>
        public static ExtendedGUID XORExtendedGUID(ExtendedGUID exGuid1, ExtendedGUID exGuid2)
        {
            byte[] exGuid1Buffer = exGuid1.SerializeToByteList().ToArray();
            byte[] exGuid2Buffer = exGuid2.SerializeToByteList().ToArray();
            byte[] resultBuffer = new byte[exGuid1Buffer.Length];

            for (int i = 0; i < exGuid1Buffer.Length; i++)
            {
                resultBuffer[i] = (byte)(exGuid1Buffer[i] ^ exGuid2Buffer[2]);
            }
            ExtendedGUID resultExGuid = new ExtendedGUID();
            resultExGuid.DoDeserializeFromByteArray(resultBuffer, 0);
            return resultExGuid;
        }
        #endregion Methods
    }
}
