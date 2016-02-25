namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides method FindItem defined in MS-OXWSSRCH.
    /// </summary>
    public interface IMS_OXWSSRCHAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Switch the current user to the new user, with the identity of the new role to communicate with server.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        void SwitchUser(string userName, string password, string domain);

        /// <summary>
        /// Find item on the server.
        /// </summary>
        /// <param name="findItemRequest">Find item operation request type.</param>
        /// <returns>Find item operation response type.</returns>
        FindItemResponseType FindItem(FindItemType findItemRequest);
    }
}