namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of SUT Control Adapter.
    /// </summary>
    public interface IMS_OFFICIALFILESUTControlAdapter : IAdapter
    {
        /// <summary>
        /// This method is used to remove all the submitted files in the specified site and list library.
        /// </summary>
        /// <param name="siteUrl">The URL of the site which the submitted files will be deleted.</param>
        /// <param name="listName">The name of the list library in which the submitted files will be deleted.</param>
        /// <returns>Return true if remove operation succeeds, otherwise return false.</returns>
        [MethodHelp("Remove all the submitted files in the specified site (siteUrl) and the specified list library (listName). If the files are removed successfully, then enter true, otherwise enter false.")]
        bool DeleteAllFiles(string siteUrl, string listName);

        /// <summary>
        /// This method is used to un-hold all the items of the specified list library from the specified hold.
        /// </summary>
        /// <param name="siteUrl">The URL of the site which the items will needs un-hold.</param>
        /// <param name="holdName">The name of the hold which hold the items.</param>
        /// <param name="listName">The name of the list library in which the items needs un-hold.</param>
        /// <returns>Return true if un-hold operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"Remove the hold on all the items in the specified list or library (listName) from the specified hold (holdName) in the site (siteUrl).
                    If succeeds then enter true, otherwise enter false.")]
        bool UnholdFiles(string siteUrl, string holdName, string listName);

        /// <summary>
        /// This method is used to switch on/off the file metadata parsing feature for the specified site.
        /// </summary>
        /// <param name="siteUrl">The URL of the site in which the file metadata parsing feature is switch on/off.</param>
        /// <param name="isEnable">The value specifies whether enable the file metadata parsing feature. If true then enable, otherwise disable.</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        [MethodHelp(@"If the specified parameter (isEnable) is true, then change the status of the site (siteUrl) that enables the file metadata parsing.
                    Otherwise disable the file metadata parsing. If the operation succeeds, enter true, otherwise enter false.")]
        bool SwitchFileMetaDataParsingFeature(string siteUrl, bool isEnable);
    }
}