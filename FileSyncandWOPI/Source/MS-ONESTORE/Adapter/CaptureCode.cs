//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_ONESTORE
{
    using System.Collections.Generic;
    using TestTools;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_ONESTOREAdapter
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="site"></param>
        public void VerifyRevisionStoreFile(OneNoteRevisionStoreFile file, ITestSite site)
        {
            this.VerifyHeader(file.Header, site);
            this.VerifyFreeChunkList(file.FreeChunkList, site);
            this.VerifyTransactionLog(file.TransactionLog, site);
            this.VerifyHashedChunkList(file.HashedChunkList, site);
            this.VerifyFileNodeList(file.FileNodeList, site);
        }

        private void VerifyHeader(Header header, ITestSite site)
        {

        }

        private void VerifyFreeChunkList(List<FreeChunkListFragment> freeChunkList, ITestSite site)
        {

        }

        private void VerifyTransactionLog(List<TransactionLogFragment> transactionLog, ITestSite site)
        {

        }

        private void VerifyHashedChunkList(List<FileNodeListFragment> hashedChunkList, ITestSite site)
        {

        }

        private void VerifyFileNodeList(List<FileNodeListFragment> fileNodeList, ITestSite site)
        {

        }
    }
}
