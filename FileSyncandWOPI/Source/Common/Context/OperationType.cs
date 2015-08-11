namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The operation type is used to distinguish the operations in MS-WOPI or MS-FSSHTTP.
    /// </summary>
    public enum OperationType
    {
        /// <summary>
        /// The operations defined in the MS-FSSHTTP.
        /// </summary>
        FSSHTTPCellStorageRequest,

        /// <summary>
        /// The ExecuteCellStorageRequest defined in the MS-WOPI.
        /// </summary>
        WOPICellStorageRequest,

        /// <summary>
        /// The ExecuteCellStorageRelativeRequest defined in the MS-WOPI. 
        /// </summary>
        WOPICellStorageRelativeRequest
    }
}