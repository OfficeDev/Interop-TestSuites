namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System;

    /// <summary>
    /// The class provides the methods to assist MS-XXXXAdapter.
    /// </summary>
    public class AdapterHelper
    {
        #region Adapter Help Methods
        public static Guid ReadGuid(byte[] byteArray, int startIndex)
        {
            byte[] guidBuffer = new byte[16];
            Array.Copy(byteArray, startIndex, guidBuffer, 0, 16);

            return new Guid(guidBuffer);
        }
        #endregion Methods
    }
}
