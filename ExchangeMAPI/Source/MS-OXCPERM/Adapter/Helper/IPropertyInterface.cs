namespace Microsoft.Protocols.TestSuites.MS_OXCPERM
{
    /// <summary>
    /// This interface defines a Base structure used by others
    /// </summary>
    public interface IPropertyInterface
    {
        /// <summary>
        /// Input the value from rawData to structure
        /// </summary>
        /// <param name="rawData">The byte array returned from the GetLists</param>
        /// <param name="count">The count point to the current digit</param>
        /// <returns>A IPropertyInterface structure contains the value</returns>
        IPropertyInterface InputValue(byte[] rawData, ref int count);

        /// <summary>
        /// Output the value saved in the IPropertyInterface structure
        /// </summary>
        /// <param name="list">The list from the input</param>
        /// <returns>Certain value of each property</returns>
        object OutputValue(IPropertyInterface list);
    }
}