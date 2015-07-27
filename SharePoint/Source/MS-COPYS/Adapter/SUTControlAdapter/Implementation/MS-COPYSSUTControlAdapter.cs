//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the SUT control managed code adapter interface.
    /// </summary>
    public class MS_COPYSSUTControlAdapter : ManagedAdapterBase, IMS_COPYSSUTControlAdapter
    {   
        /// <summary>
        /// Represents the error messages generated in delete files' process.
        /// </summary>
        private StringBuilder errorMessageInDeleteProcess = new StringBuilder();

        /// <summary>
        /// Represents the MS-LISTSWS proxy instance which is used to invoke functions of Lists web service.
        /// </summary>
        private ListsSoap listswsProxy = null;

        /// <summary>
        /// Initialize the adapter instance.
        /// </summary>
        /// <param name="testSite">A return value represents the ITestSite instance which contains the test context.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            TestSuiteManageHelper.Initialize(this.Site);
            TestSuiteManageHelper.AcceptServerCertificate();

            // Initial the listsws proxy without schema validation.
            if (null == this.listswsProxy)
            {
                this.listswsProxy = Proxy.CreateProxy<ListsSoap>(testSite, true, false, false);
            }

            FileUrlHelper.ValidateFileUrl(TestSuiteManageHelper.TargetServiceUrlOfMSCOPYS);

            // Point to listsws service according to the MS-COPYS service URL.
            string targetServiceUrl = Path.GetDirectoryName(TestSuiteManageHelper.TargetServiceUrlOfMSCOPYS);
            targetServiceUrl = Path.Combine(targetServiceUrl, @"lists.asmx");

            // Work around for local path format mapping to URL path format.
            targetServiceUrl = targetServiceUrl.Replace(@"\", @"/");
            targetServiceUrl = targetServiceUrl.Replace(@":/", @"://");

            // Setting the properties of listsws service proxy.
            this.listswsProxy.Url = targetServiceUrl;
            this.listswsProxy.Credentials = TestSuiteManageHelper.DefaultUserCredential;
            this.listswsProxy.SoapVersion = TestSuiteManageHelper.GetSoapProtoclVersionByCurrentSetting();

            // 60000 means the configure SOAP Timeout is in minute.
            this.listswsProxy.Timeout = TestSuiteManageHelper.CurrentSoapTimeOutValue;
        }

        /// <summary>
        /// This method is used to remove the files.
        /// </summary>
        /// <param name="fileUrls">Specify the file URLs that will be removed. Each file URL is split by ";" symbol</param>
        /// <returns>Return true if the operation succeeds, otherwise return false.</returns>
        public bool DeleteFiles(string fileUrls)
        {
            if (string.IsNullOrEmpty(fileUrls))
            {
                throw new ArgumentNullException("fileUrls");
            }

            string[] fileUrlCollection = fileUrls.Split(new string[] { @";" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string fileUrlItem in fileUrlCollection)
            {
                Uri filePath;
                if (!Uri.TryCreate(fileUrlItem, UriKind.Absolute, out filePath))
                {
                    string errorMsg = string.Format(@"The file URL item[{0}] is not a valid URL. Ignore this.");
                    this.errorMessageInDeleteProcess.AppendLine(errorMsg);
                    continue;
                }

                this.DeleteSingleFile(fileUrlItem);
            }

            if (this.errorMessageInDeleteProcess.Length > 0)
            {   
                this.Site.Log.Add(
                                LogEntryKind.Debug, 
                                "There are some errors generated in the delete file process:\r\n[{0}]",
                                this.errorMessageInDeleteProcess.ToString());
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// This method is used to upload a file to the specified full file URL. The file's content will be random generated, and encoded with UTF8.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file, where the file will be uploaded.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        public bool UploadTextFile(string fileUrl)
        {
            WebClient client = new WebClient();
            try
            {
                byte[] contents = System.Text.Encoding.UTF8.GetBytes(Common.GenerateResourceName(Site, "FileContent"));
                client.Credentials = TestSuiteManageHelper.DefaultUserCredential;
                client.UploadData(fileUrl, "PUT", contents);
            }
            catch (System.Net.WebException ex)
            {
                Site.Log.Add(
                    LogEntryKind.TestError,
                    string.Format("Cannot upload the file to the full URL {0}, the exception message is {1}", fileUrl, ex.Message));

                return false;
            }
            finally
            {
                if (client != null)
                {
                    client.Dispose();
                }
            }

            return true;
        }

        /// <summary>
        /// A method used to check out a file by specified user credential.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file which will be check out by specified user.</param>
        /// <param name="userName">A parameter represents the user name which will undo checkout the file. The file must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        public bool CheckOutFileByUser(string fileUrl, string userName, string password, string domain)
        { 
            #region parameter validation
            FileUrlHelper.ValidateFileUrl(fileUrl);

            if (string.IsNullOrEmpty(userName))
            {
              throw new ArgumentNullException("userName");
            }

            #endregion parameter validation
            
            if (null == this.listswsProxy)
            {
                throw new InvalidOperationException("The LISTSWS proxy is not initialized, should call the [Initialize] method before calling this method.");
            }

            this.listswsProxy.Credentials = new NetworkCredential(userName, password, domain);

            bool checkOutResult;
            try
            {
                checkOutResult = this.listswsProxy.CheckOutFile(fileUrl, bool.TrueString, string.Empty);
            }
            catch (SoapException soapEx)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when the SUT control adapter try to check out a file[{0}]:\r\nExcetption Message:\r\n[{1}]\r\\nStackTrace:\r\n[{2}]",
                                fileUrl,
                                string.IsNullOrEmpty(soapEx.Message) ? "None" : soapEx.Message,
                                string.IsNullOrEmpty(soapEx.StackTrace) ? "None" : soapEx.StackTrace);
                return false;
            }
            finally
            {
                this.listswsProxy.Credentials = TestSuiteManageHelper.DefaultUserCredential;
            }

            return checkOutResult;
        }

        /// <summary>
        /// A method used to undo checkout for a file by specified user credential.
        /// </summary>
        /// <param name="fileUrl">A parameter represents the absolute URL of a file which will be undo checkout by specified user.</param>
        /// <param name="userName">A parameter represents the user name which will check out the file. The file must be stored in the destination SUT which is indicated by "SutComputerName" property in "SharePointCommonConfiguration.deployment.ptfconfig" file.</param>
        /// <param name="password">A parameter represents the password of the user.</param>
        /// <param name="domain">A parameter represents the domain of the user.</param>
        /// <returns>Return true if the operation succeeds, otherwise returns false.</returns>
        public bool UndoCheckOutFileByUser(string fileUrl, string userName, string password, string domain)
        {
            #region parameter validation
            FileUrlHelper.ValidateFileUrl(fileUrl);

            if (string.IsNullOrEmpty(userName))
            {
                throw new ArgumentNullException("userName");
            }

            #endregion parameter validation

            if (null == this.listswsProxy)
            {
                throw new InvalidOperationException("The LISTSWS proxy is not initialized, should call the [Initialize] method before calling this method.");
            }

            this.listswsProxy.Credentials = new NetworkCredential(userName, password, domain);

            bool undoCheckOutResult;
            try
            {
                undoCheckOutResult = this.listswsProxy.UndoCheckOut(fileUrl);
            }
            catch (SoapException soapEx)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is an exception generated when the SUT control adapter try to undo check out for a file[{0}]:\r\nExcetption Message:\r\n[{1}]\r\\nStackTrace:\r\n[{2}]",
                                fileUrl,
                                string.IsNullOrEmpty(soapEx.Message) ? "None" : soapEx.Message,
                                string.IsNullOrEmpty(soapEx.StackTrace) ? "None" : soapEx.StackTrace);
                return false;
            }
            finally
            {
                this.listswsProxy.Credentials = TestSuiteManageHelper.DefaultUserCredential;
            }

            return undoCheckOutResult;
        }
 
        /// <summary>
        /// This method is used to remove the file.
        /// </summary>
        /// <param name="singleFileUrl">Specify the file URL that will be removed.</param>
        private void DeleteSingleFile(string singleFileUrl)
        {
            HttpWebRequest deleteRequest = HttpWebRequest.Create(singleFileUrl) as HttpWebRequest;
            HttpWebResponse response = null;

            deleteRequest.Credentials = TestSuiteManageHelper.DefaultUserCredential;
            deleteRequest.Method = "DELETE";

            try
            {
                response = deleteRequest.GetResponse() as HttpWebResponse;
            }
            catch (System.Net.WebException ex)
            {
                string errorMsg = string.Format(
                                                @"Cannot delete the file[{0}], the exception message is {1}",
                                                singleFileUrl,
                                                ex.Message);
                this.errorMessageInDeleteProcess.AppendLine(errorMsg);
            }
            finally
            {
                if (response != null)
                {
                    response.Close();
                }
            }

            // For the loop, sleep zero to ensure other control thread can visit the CPU resource. This is for optimizing performance.
            System.Threading.Thread.Sleep(0);
        }
    }
}
