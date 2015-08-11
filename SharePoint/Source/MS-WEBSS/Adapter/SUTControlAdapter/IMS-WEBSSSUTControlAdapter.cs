namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The Interface is used to configure the SUT.
    /// </summary>
    public interface IMS_WEBSSSUTControlAdapter : IAdapter
    {
        #region Common scripts

        /// <summary>
        /// Used to set Read-Only or Sealed properties of the ContentType.
        /// </summary>
        /// <param name="webSiteName">Specify the web the list or listItem or active features are located.</param>
        /// <param name="contentTypeName">Specify the site contentType name.</param>
        /// <param name="isReadOnly">set contentType is Read-Only or not.</param>
        /// <param name="isSealed">set contentType is sealed or not.</param>
        [MethodHelp("Based on the action parameters, set the content type property of \"contentTypeName\" to \"ReadOnly\" or \"Sealed\"  under the specified website."
                    + "If the action parameter \"isReadOnly\" is set to true, then set the \"ReadOnly\" property to true; if \"isReadyOnly\" is set tofalse, then set the \"ReadOnly\" property to false."
                    + "If the Action parameter \"isSealed\" is set to true, then set the \"Sealed\" property to true; if \"isSealed\" is set to false, then set the \"Sealed\" property to false.")]
        void SetContentTypeReadOnlyOrSealed(string webSiteName, string contentTypeName, bool isReadOnly, bool isSealed);

        /// <summary>
        /// Get the Object Id(s) from list or listItem or active feature of the web site.
        /// </summary>
        /// <param name="webSiteName">Specify the web the list or listItem or active features are located.</param>
        /// <param name="objectName">Must be one of "list", "listItem", "site_features" and "site_collection_features".</param>
        /// <returns>The Object ID of list or listItem or Feature.</returns>
        [MethodHelp("Based on the \"WebsiteName\" and specified \"ObjectName\", this method is used to get the Object Id(s) Of the default list, or the default listItem, or the active features."
                    + "If the action parameter \"webSiteName\" is not a valid website, then enter \"null\" in the \"return value\" field."
                    + "If the action parameter \"webSiteName\" is valid and the action parameter \"objectName\" is set to \"list\", then in the \"return value\" field, enter the \"ID\" of the default list on the website."
                    + "If the action parameter \"webSiteName\" is valid and the action parameter \"objectName\" is set to \"listItem\", then in the \"return value\" field, enter the \"ID\" of the default listItem in the default list on the website."
                    + "If the action parameter \"webSiteName\" is valid and the action parameter \"objectName\" is set to \"site_features\" or \"site_collection_features\", then in the \"return value\" field, enter one or more ID(s) of active features (separated by one space) on the website."
                    + "If the action parameter \"webSiteName\" is valid and the action parameter \"objectName\" is not set to the above values, then enter \"null\" in the \"return value\" field.")]
        string GetObjectId(string webSiteName, string objectName);

        #endregion
    }
}