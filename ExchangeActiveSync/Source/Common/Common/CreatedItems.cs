namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// The items created by case.
    /// </summary>
    public class CreatedItems
    {
        /// <summary>
        /// Define a collection used to store the subject of created items.
        /// </summary>
        private Collection<string> itemSubject;

        /// <summary>
        /// Initializes a new instance of the CreatedItems class.
        /// </summary>
        public CreatedItems()
        {
            // Initialize the itemSubject collection.
            if (this.itemSubject != null)
            {
                this.itemSubject.Clear();
            }
            else
            {
                this.itemSubject = new Collection<string>();
            }
        }

        /// <summary>
        /// Gets or sets the CollectionId of the created item's parent folder.
        /// </summary>
        public string CollectionId { get; set; }

        /// <summary>
        /// Gets the subjects of the created items.
        /// </summary>
        public Collection<string> ItemSubject
        {
            get { return this.itemSubject; }
        }
    }
}