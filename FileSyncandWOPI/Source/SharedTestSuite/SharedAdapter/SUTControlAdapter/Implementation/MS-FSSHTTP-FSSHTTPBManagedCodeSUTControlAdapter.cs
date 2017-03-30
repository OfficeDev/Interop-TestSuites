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