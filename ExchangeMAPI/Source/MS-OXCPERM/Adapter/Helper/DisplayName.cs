namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    /// <summary>
    /// This structure represents the input and output method and the value type of the name of the address list that is contained in the OAB message
    /// </summary>
    public struct DisplayName : IPropertyInterface
    {
        /// <summary>
        /// The value of this property is the name of the address list that is contained in the OAB message
        /// </summary>
        private string displayName;

        /// <summary>
        /// Input the value from rawData to structure
        /// </summary>
        /// <param name="rawData">The byte array returned from the GetLists</param>
        /// <param name="count">The count point to the current digit</param>
        /// <returns>A IPropertyInterface structure contains the value</returns>
        public IPropertyInterface InputValue(byte[] rawData, ref int count)
        {
            DisplayName value = new DisplayName();
            for (int i = 0; i <= (rawData.Length - count) / 2; i++)
            {
                if (rawData[count + (i * 2)] == 0x00)
                {
                    int stopdigit = count + (i * 2);
                    value.displayName = this.ConvertBytesToString(rawData, count, stopdigit - 1);
                    count = stopdigit + 2;
                    return value;
                }
            }

            return null;
        }

        /// <summary>
        /// Output the value saved in the IPropertyInterface structure
        /// </summary>
        /// <param name="list">The list from the input</param>
        /// <returns>Certain value of each property</returns>
        public object OutputValue(IPropertyInterface list)
        {
            DisplayName displayname = (DisplayName)list;
            string value = displayname.displayName;
            return value;
        }

        /// <summary>
        /// Convert the input bytes to a string
        /// </summary>
        /// <param name="source">Input byte array</param>
        /// <param name="start">The start digit</param>
        /// <param name="end">The end digit</param>
        /// <returns>The result string</returns>
        public string ConvertBytesToString(byte[] source, int start, int end)
        {
            int pos = start;
            string result = string.Empty;
            while (pos < end)
            {
                if (source[pos] != 0)
                {
                    result += (char)source[pos];
                }

                pos++;
            }

            return result;
        }
    }
}