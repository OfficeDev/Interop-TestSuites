namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Net;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the SUT control managed code adapter interface.
    /// </summary>
    public class MS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter : ManagedAdapterBase, IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter
    {
        /// <summary>
        /// This method is used to check in the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked in.</param>
        /// <param name="userName">Specify the name of the user who checks in the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <param name="checkInComments">Specify the checked in comments.</param>
        /// <returns>Return true if the check in succeeds, otherwise return false.</returns>
        public bool CheckInFile(string fileUrl, string userName, string password, string domain, string checkInComments)
        {
            string targetSiteCollectionUrl = Common.Common.GetConfigurationPropertyValue("TargetSiteCollectionUrl", this.Site);
            string fullFileUri = string.Format("{0}/_vti_bin/lists.asmx", targetSiteCollectionUrl);
            if (fullFileUri.StartsWith("HTTPS", System.StringComparison.OrdinalIgnoreCase))
            {
                Common.Common.AcceptServerCertificate();
            }
            ListsSoap listsProxy = new ListsSoap();
            listsProxy.Url = fullFileUri;
            listsProxy.Credentials = new NetworkCredential(userName, password, domain);

            return listsProxy.CheckInFile(fileUrl, checkInComments, "1");
        }

        /// <summary>
        /// This method is used to check out the specified file using the specified credential.
        /// </summary>
        /// <param name="fileUrl">Specify the absolute URL of a file which needs to be checked out.</param>
        /// <param name="userName">Specify the name of the user who checks out the file.</param>
        /// <param name="password">Specify the password of the user.</param>
        /// <param name="domain">Specify the domain of the user.</param>
        /// <returns>Return true if the check out succeeds, otherwise return false.</returns>
        public bool CheckOutFile(string fileUrl,string userName,string password,string domain)
        {
            string targetSiteCollectionUrl = Common.Common.GetConfigurationPropertyValue("TargetSiteCollectionUrl", this.Site);
            string fullFileUri = string.Format("{0}/_vti_bin/lists.asmx", targetSiteCollectionUrl);
            if (fullFileUri.StartsWith("HTTPS", System.StringComparison.OrdinalIgnoreCase))
            {
                Common.Common.AcceptServerCertificate();
            }
            ListsSoap listsProxy = new ListsSoap();
            listsProxy.Url = fullFileUri;
            listsProxy.Credentials = new NetworkCredential(userName, password, domain);
            return listsProxy.CheckOutFile(fileUrl, "False", null);
        }

        /// <summary>
        /// This method is used to remove the file from the path of file URI.
        /// </summary>
        /// <param name="fileUrl">Specify the URL in where the file will be removed.</param>
        /// <param name="fileName">Specify the name for the file that will be removed.</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        public bool RemoveFile(string fileUrl, string fileName)
        {
            string fullFileUri = string.Format("{0}/{1}", fileUrl, fileName);
            HttpWebRequest deleteRequest = HttpWebRequest.Create(fullFileUri) as HttpWebRequest;
            HttpWebResponse response = null;

            try
            {
                if (fullFileUri.StartsWith("HTTPS", System.StringComparison.OrdinalIgnoreCase))
                {
                    Common.Common.AcceptServerCertificate();
                }

                deleteRequest.Credentials = new NetworkCredential(Common.Common.GetConfigurationPropertyValue("UserName1", Site), Common.Common.GetConfigurationPropertyValue("Password1", Site), Common.Common.GetConfigurationPropertyValue("Domain", Site));
                deleteRequest.Method = "DELETE";

                response = deleteRequest.GetResponse() as HttpWebResponse;

                return response.StatusCode == HttpStatusCode.NoContent || response.StatusCode == HttpStatusCode.OK;
            }
            catch (System.Net.WebException ex)
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    string.Format("Cannot delete the file in the full URI {0}, the exception message is {1}", fullFileUri, ex.Message));

                return false;
            }
            finally
            {
                // Close the connection before returning.
                if (response != null)
                {
                    response.Close();
                }
            }
        }
    }
}