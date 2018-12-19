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
    using System;
    using System.Net;
    using System.Xml;
    using Microsoft.Protocols.TestTools;
    using System.IO;

    /// <summary>
    /// Adapter class of MS-ONESTORE.
    /// </summary>
    public partial class MS_ONESTOREAdapter : ManagedAdapterBase, IMS_ONESTOREAdapter
    {
        #region Variables

        #endregion Variables

        #region Initialize TestSuite

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ONESTORE";
        }

        #endregion Initialize TestSuite

        #region MS_ONESTOREAdapter Members
        /// <summary>
        /// Load and parse the OneNote revision-based file.
        /// </summary>
        /// <returns>Return the instacne of OneNoteRevisionStoreFile.</returns>
        public OneNoteRevisionStoreFile LoadOneNoteFile(string fileName)
        {
            byte[] buffer = File.ReadAllBytes(fileName);
            OneNoteRevisionStoreFile oneNoteFile = new OneNoteRevisionStoreFile();
            oneNoteFile.DoDeserializeFromByteArray(buffer);

            this.VerifyRevisionStoreFile(oneNoteFile);

            return oneNoteFile;
        }
        #endregion MS_ONESTOREAdapter Members
    }
}
