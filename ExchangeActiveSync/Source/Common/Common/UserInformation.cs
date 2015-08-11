namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.ObjectModel;

    /// <summary>
    /// The information for user
    /// </summary>
    public class UserInformation
    {
        /// <summary>
        /// Initializes a new instance of the UserInformation class.
        /// </summary>
        public UserInformation()
        {
            this.UserCreatedItems = new Collection<CreatedItems>();
            this.UserCreatedFolders = new Collection<string>();
        }

        /// <summary>
        /// Gets or sets the name of user.
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the password of user
        /// </summary>
        public string UserPassword { get; set; }

        /// <summary>
        /// Gets or sets the name of domain
        /// </summary>
        public string UserDomain { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Inbox folder
        /// </summary>
        public string InboxCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Sent Items folder
        /// </summary>
        public string SentItemsCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Deleted Items folder
        /// </summary>
        public string DeletedItemsCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the Recipient information cache
        /// </summary>
        public string RecipientInformationCacheCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Calendar folder
        /// </summary>
        public string CalendarCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Contacts folder
        /// </summary>
        public string ContactsCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Notes folder
        /// </summary>
        public string NotesCollectionId { get; set; }

        /// <summary>
        /// Gets or sets the server ID of the default Tasks folder
        /// </summary>
        public string TasksCollectionId { get; set; }

        /// <summary>
        /// Gets the items created by the user
        /// </summary>
        public Collection<CreatedItems> UserCreatedItems { get; private set; }

        /// <summary>
        /// Gets the folders created by the user
        /// </summary>
        public Collection<string> UserCreatedFolders { get; private set; }
    }
}