namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// A class is used to support WOPI resource URL cache function.
    /// </summary>
    internal class WOPIResourceUrlCache
    {   
        /// <summary>
        /// A Dictionary type instance which is used to store cache information for specified user.
        /// </summary>
        private static Dictionary<string, Dictionary<string, string>> wopiResourceUrlCaches = new Dictionary<string, Dictionary<string, string>>();
        
        /// <summary>
        /// A method is used to add a  WOPI resource URL mapping record into the cache. The record is determined by the domain\user and the file url.
        /// </summary>
        /// <param name="userName">A parameter represents the user name.</param>
        /// <param name="domain">A parameter represents the domain name.</param>
        /// <param name="absoluteFileUrl">A parameter represents the file URL. It must be absolute URL.</param>
        /// <param name="wopiResourceUrl">A parameter represents the WOPI resource URL.</param>
        public static void AddWOPIResourceUrl(string userName, string domain, string absoluteFileUrl, string wopiResourceUrl)
        {
            #region Verify parameter 
            HelperBase.CheckInputParameterNullOrEmpty<string>(userName, "userName", "AddWOPIResourceUrl");
            HelperBase.CheckInputParameterNullOrEmpty<string>(domain, "domain", "AddWOPIResourceUrl");
            HelperBase.CheckInputParameterNullOrEmpty<string>(absoluteFileUrl, "absoluteFileUrl", "AddWOPIResourceUrl");
            HelperBase.CheckInputParameterNullOrEmpty<string>(wopiResourceUrl, "wopiResourceUrl", "AddWOPIResourceUrl");
            
            Uri filelocation;
            if (!Uri.TryCreate(absoluteFileUrl, UriKind.Absolute, out filelocation))
            {
                throw new ArgumentException("The file url should be absolute url", "absoluteFileUrl");
            }

            Uri wopiResourceLocation;
            if (!Uri.TryCreate(wopiResourceUrl, UriKind.Absolute, out wopiResourceLocation))
            {
                throw new ArgumentException("The WOPI resource url should be absolute url", "wopiResourceUrl");
            }

            #endregion 

            string fullUserName = string.Format(@"{0}\{1}", domain, userName);
            Dictionary<string, string> urlMappingCollectionOfUser = GetUrlMappingCollectionByFullUserName(fullUserName);
            urlMappingCollectionOfUser.Add(absoluteFileUrl, wopiResourceUrl);
        }
 
        /// <summary>
        /// A method is used to verify whether a WOPI resource URL mapping record exists in cache. The WOPI resource url record is determined by the domain\user and the file url.
        /// </summary>
        /// <param name="userName">A parameter represents the user name.</param>
        /// <param name="domain">A parameter represents the domain name.</param>
        /// <param name="absoluteFileUrl">A parameter represents the file URL. It must be absolute URL.</param>
        /// <param name="wopiResourceUrl">An out parameter represents the WOPI resource URL when this method returns. If this method could not get the WOPI resource URL, this value will be string.Empty.</param>
        /// <returns>Return 'true' indicating the WOPI resource URL record exists in cache and out put to the "wopiResourceUrl" parameter.</returns>
        public static bool TryGetWOPIResourceUrl(string userName, string domain, string absoluteFileUrl, out string wopiResourceUrl)
        {
            #region Verify parameter
            HelperBase.CheckInputParameterNullOrEmpty<string>(userName, "userName", "TryGetWOPIResourceUrl");
            HelperBase.CheckInputParameterNullOrEmpty<string>(domain, "domain", "TryGetWOPIResourceUrl");
            HelperBase.CheckInputParameterNullOrEmpty<string>(absoluteFileUrl, "absoluteFileUrl", "TryGetWOPIResourceUrl");
   
            Uri fileLocation;
            if (!Uri.TryCreate(absoluteFileUrl, UriKind.Absolute, out fileLocation))
            {
                throw new ArgumentException("The file url should be absolute url", "absoluteFileUrl");
            }

            #endregion 

            string fullUserName = string.Format(@"{0}\{1}", domain, userName);
            Dictionary<string, string> urlMappingCollectionOfUser = GetUrlMappingCollectionByFullUserName(fullUserName, false);
            wopiResourceUrl = string.Empty;
            if (null == urlMappingCollectionOfUser || 0 == urlMappingCollectionOfUser.Count)
            {
                return false;
            }

            var expectedWOPIResourceUrlItems = from wopiUrlMapping in urlMappingCollectionOfUser
                                               where wopiUrlMapping.Key.Equals(absoluteFileUrl, StringComparison.OrdinalIgnoreCase)
                                               select wopiUrlMapping.Value;
            
            if (0 == expectedWOPIResourceUrlItems.Count())
            {
                return false;
            }
 
            wopiResourceUrl = expectedWOPIResourceUrlItems.ElementAt<string>(0);
            return true;
        }

        /// <summary>
        /// A method is used to get URL mapping collection by specified full user name. If the cache dose not include the specified user, this method will append a new record to catch, and return the new mapping collection which is used to store the mapping between the file URL and WOPI URL.
        /// </summary>
        /// <param name="fullUserName">A parameter represents the full user name, its format is domain\username.</param>
        /// <param name="isCreateNewForNotExistUser">A parameter represents a bool value indicating whether create new mapping collection if the cache does not contain the specified user. 'true' means the method will create new mapping collection, default value is 'true'</param>
        /// <returns>A return value represents the URL mapping collection which is used to store the mapping between file URL and the WOPI resource URL for a user.</returns>
        protected static Dictionary<string, string> GetUrlMappingCollectionByFullUserName(string fullUserName, bool isCreateNewForNotExistUser = true)
        {
            var expectedMappingCollection = from wopiUrlRecord in wopiResourceUrlCaches
                                            where wopiUrlRecord.Key.Equals(fullUserName, StringComparison.OrdinalIgnoreCase)
                                            select wopiUrlRecord.Value;

            // If the user is not in cache, add a new record and return the mapping collection of the record.
            if (0 == expectedMappingCollection.Count())
            {
                Dictionary<string, string> urlMappingCollectionItem = null;
                if (isCreateNewForNotExistUser)
                {
                    urlMappingCollectionItem = new Dictionary<string, string>();
                    wopiResourceUrlCaches.Add(fullUserName, urlMappingCollectionItem);
                }

                return urlMappingCollectionItem;
            }
            else
            {   
                return expectedMappingCollection.ElementAt<Dictionary<string, string>>(0);           
            }
        } 
    }
}