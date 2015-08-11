namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Contain the information to be changed for calendar related item.
    /// </summary>
    public class AdapterHelper
    {
        /// <summary>
        /// Gets or sets the identifier of item to be updated.
        /// </summary>
        public BaseItemIdType ItemId { get; set; }

        /// <summary>
        /// Gets or sets the URIs of well-known element to be updated.
        /// </summary>
        public UnindexedFieldURIType FieldURI { get; set; }

        /// <summary>
        /// Gets or sets the item used to store the element to be updated.
        /// </summary>
        public ItemType Item { get; set; }
    }
}