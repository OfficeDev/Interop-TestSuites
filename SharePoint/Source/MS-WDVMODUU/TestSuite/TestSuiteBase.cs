//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WDVMODUU
{
    using System;
    using System.Collections;
    using System.Collections.Specialized;
    using System.IO;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
        
    /// <summary>
    /// The base class of other test suite classes, it includes common methods and properties that used by its child test suite classes.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Private Variables

        /// <summary>
        /// The instance of interface IMS_WDVMODUUAdapter.
        /// </summary>
        private IMS_WDVMODUUAdapter wdvmoduuAdapter;

        /// <summary>
        /// The array list of file URI that should be deleted in test clean up method.
        /// </summary>
        private ArrayList arrayListOfDeleteFile;

        #endregion  Private Variables

        #region Protected Properties that used by child test suite class

        /// <summary>
        /// Gets the protocol adapter object.
        /// </summary>
        protected IMS_WDVMODUUAdapter Adapter
        {
            get
            {
                return this.wdvmoduuAdapter;
            }
        }

        /// <summary>
        /// Gets the array list for file URI that will be deleted in test clean up.
        /// </summary>
        protected ArrayList ArrayListForDeleteFile
        {
            get
            {
                return this.arrayListOfDeleteFile;
            }
        }
        
        #endregion Protected Properties that used by child test suite class

        #region TestInitialize method

        /// <summary>
        /// A test case's level initialization method for TestSuiteBase class. It will perform before each test case.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.wdvmoduuAdapter = this.Site.GetAdapter<IMS_WDVMODUUAdapter>();
            Common.CheckCommonProperties(this.Site, false);

            if (TransportProtocol.HTTPS == Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site))
            {
                this.Site.Assume.Inconclusive("The open specification [MS-WDVMODUU] does not support to run under HTTPS transport.");
            }

            this.arrayListOfDeleteFile = new ArrayList();
        }

        #endregion

        #region TestCleanup method

        /// <summary>
        /// This method will run after test case executes.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            // Delete all files in the internal delete list from the server.
            IEnumerator enumeratorDeleteList = this.ArrayListForDeleteFile.GetEnumerator();
            while (enumeratorDeleteList.MoveNext())
            {
                string currentFileUri = (string)enumeratorDeleteList.Current;
                bool isSuccessful = this.DeleteTheFileInTheServer(currentFileUri);
                Site.Assert.IsTrue(
                    isSuccessful,
                    string.Format("Failed to delete file {0} in 'TestCleanup'.", currentFileUri));
            }

            this.arrayListOfDeleteFile.Clear();
        }

        #endregion

        #region Protected Methods that used by child test suite class
 
        /// <summary>
        /// Remove the file URI from the internal array list "arrayListOfDeleteFile".
        /// </summary>
        /// <param name="fileUri">The file URI that need to be removed.</param>
        protected void RemoveFileUriFromDeleteList(string fileUri)
        {
            IEnumerator enumeratorDeleteList = this.ArrayListForDeleteFile.GetEnumerator();
            while (enumeratorDeleteList.MoveNext())
            {
                string currentFileUri = (string)enumeratorDeleteList.Current;
                if (string.Compare(fileUri, currentFileUri, true) == 0)
                {
                    this.ArrayListForDeleteFile.Remove(enumeratorDeleteList.Current);
                    enumeratorDeleteList = this.ArrayListForDeleteFile.GetEnumerator();
                }
            }
        }

        /// <summary>
        /// Delete the file in the internal array list "arrayListOfDeleteFile" from the server.
        /// </summary>
        /// <param name="destinationUri">The file URI that will be deleted from the server.</param>
        /// <returns>Return true if delete the file successfully, else return false.</returns>
        protected bool DeleteTheFileInTheServer(string destinationUri)
        {
            if (string.IsNullOrEmpty(destinationUri))
            {
                return false;
            }

            bool isSuccessful = false;
            
            // Construct the request headers.
            NameValueCollection headersCollection = new NameValueCollection();
            headersCollection.Clear();
            headersCollection.Add("ProtocolVersion", "HTTP/1.1");

            // Send HTTP DELETE method to delete the file.
            WDVMODUUResponse response = this.Adapter.Delete(destinationUri, headersCollection);

            // Assert the DELETE method is successful.
            if (response.StatusCode == HttpStatusCode.NoContent)
            {
                isSuccessful = true;
            }
            else
            {
                isSuccessful = false;
                string errorInfo = string.Format(
                    "Fail to delete file {0} in the server! The return status code is '{1}', and the return status description is '{2}'.", 
                    destinationUri, 
                    (int)response.StatusCode, 
                    response.StatusDescription);
                this.Site.Assert.Fail(errorInfo);
            }

            return isSuccessful;
        }

        #region GetLocalFileContent method

        /// <summary>
        /// Retrieve a byte array contains the content of the given file.
        /// </summary>
        /// <param name="fileName">The name of file whose content is to be retrieved.</param>
        /// <returns>The byte array contains the content of the given file.</returns>
        protected byte[] GetLocalFileContent(string fileName)
        {
            byte[] bytes = null;

            // Assert the file is existed in the local folder.
            Site.Assume.IsTrue(File.Exists(fileName), "The file '{0}' should exist in local output folder!", fileName);

            // Get the bytes array contains the content of the given file.
            using (FileStream fileStream = File.OpenRead(fileName))
            {
                bytes = new byte[fileStream.Length];
                fileStream.Read(bytes, 0, Convert.ToInt32(fileStream.Length));
                fileStream.Close();
            }

            return bytes;
        }

        #endregion

        #endregion Protected Methods that used by child test suite class
    }
}