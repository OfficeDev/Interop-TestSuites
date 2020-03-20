namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
 
    /// <summary>
    /// The adapter interface which provides methods defined in MS-OXWSMTGS.
    /// </summary>
    public interface IMS_OXWSMTGSAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Get the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the GetItem operation.</param>
        /// <returns>The response message returned by GetItem operation.</returns>
        GetItemResponseType GetItem(GetItemType request);

        /// <summary>
        /// Delete the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the DeleteItem operation.</param>
        /// <returns>The response message returned by DeleteItem operation.</returns>
        DeleteItemResponseType DeleteItem(DeleteItemType request);

        /// <summary>
        /// Update the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the UpdateItem operation.</param>
        /// <returns>The response message returned by UpdateItem operation.</returns>
        UpdateItemResponseType UpdateItem(UpdateItemType request);

        /// <summary>
        /// Move the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the MoveItem operation.</param>
        /// <returns>The response message returned by MoveItem operation.</returns>
        MoveItemResponseType MoveItem(MoveItemType request);

        /// <summary>
        /// Copy the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the CopyItem operation.</param>
        /// <returns>The response message returned by CopyItem operation.</returns>
        CopyItemResponseType CopyItem(CopyItemType request);

        /// <summary>
        /// Create the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the CreateItem operation.</param>
        /// <returns>The response message returned by CreateItem operation.</returns>
        CreateItemResponseType CreateItem(CreateItemType request);

        /// <summary>
        /// Retrieves the profile image for a mailbox
        /// </summary>
        /// <param name="getRemindersRequest">The request of GetReminders operation.</param>
        /// <returns>A response to GetReminders operation request.</returns>
        GetRemindersResponseMessageType GetReminders(GetRemindersType getRemindersRequest);

        /// <summary>
        /// Retrieves the profile image for a mailbox
        /// </summary>
        /// <param name="PerformReminderActionRequest">The request of PerformReminderAction operation.</param>
        /// <returns>A response to PerformReminderAction operation request.</returns>
        PerformReminderActionResponseMessageType PerformReminderAction(PerformReminderActionType PerformReminderActionRequest);

        /// <summary>
        /// Switch the current user to a new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}