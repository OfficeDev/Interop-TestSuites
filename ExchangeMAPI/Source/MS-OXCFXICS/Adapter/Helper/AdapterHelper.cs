//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter helper
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// Index of the server handle
        /// </summary>
        private static int handleIndex;

        /// <summary>
        /// Index of the server object id
        /// </summary>
        private static int objectIdIndex;

        /// <summary>
        /// the index of folder name
        /// </summary>
        private static int folderNameIndex;

        /// <summary>
        /// Index of the stream buffer
        /// </summary>
        private static int streamBufferIndex;

        /// <summary>
        /// index of ICS state.
        /// </summary>
        private static int icsStateIndex;

        /// <summary>
        /// index of CnsetRead.
        /// </summary>
        private static int cnsetReadIndex;

        /// <summary>
        /// index of CnsetSeen.
        /// </summary>
        private static int cnsetSeenIndex;

        /// <summary>
        /// index of CnsetSeenFAI.
        /// </summary>
        private static int cnsetSeenFAIIndex;

        /// <summary>
        /// The site of the test suite.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Gets or sets the value of the test suite site .
        /// </summary>
        public static ITestSite Site
        {
            get 
            {
                return site;
            }

            set 
            {
                site = value;
            }
        }

        /// <summary>
        /// gets a unique CnsetRead index.
        /// </summary>
        /// <returns>a unique CnsetRead index.</returns>
        public static int GetCnsetReadIndex()
        {
            return ++cnsetReadIndex;
        }

        /// <summary>
        /// gets a unique CnsetSeen index.
        /// </summary>
        /// <returns>a unique CnsetSeen index.</returns>
        public static int GetCnsetSeenIndex()
        {
            return ++cnsetSeenIndex;
        }

        /// <summary>
        /// gets a unique CnsetSeenFAI index.
        /// </summary>
        /// <returns>a unique CnsetSeenFAI index.</returns>
        public static int GetCnsetSeenFAIIndex()
        {
            return ++cnsetSeenFAIIndex;
        }

        /// <summary>
        /// Get the index of handle
        /// </summary>
        /// <returns>The index of handle</returns>
        public static int GetHandleIndex()
        {
            int tempHandleIndex = ++handleIndex;
            return tempHandleIndex;
        }

        /// <summary>
        /// gets a unique folder name index.
        /// </summary>
        /// <returns>a unique folder name index.</returns>
        public static int GetFolderNameIndex()
        {
            int tempHandleIndex = ++folderNameIndex;
            return tempHandleIndex;
        }

        /// <summary>
        /// Get index of the server object id
        /// </summary>
        /// <returns>The index of the server object id</returns>
        public static int GetObjectIdIndex()
        {
            int tempObjectIdIndex = ++objectIdIndex;
            return tempObjectIdIndex;
        }

        /// <summary>
        /// Get index of the stream buffer
        /// </summary>
        /// <returns>The index of the stream buffer</returns>
        public static int GetStreamBufferIndex()
        {
            int tempObjectIdIndex = ++streamBufferIndex;
            return tempObjectIdIndex;
        }

        /// <summary>
        /// gets a unique ICS state index.
        /// </summary>
        /// <returns>a unique ICS state index.</returns>
        public static int GetICSStateIndex()
        {
            int tempICSStateIndex = ++icsStateIndex;
            return tempICSStateIndex;
        }

        /// <summary>
        /// Reset the index
        /// </summary>
        public static void ClearIndex()
        {
            handleIndex = 0;
            objectIdIndex = 0;
            streamBufferIndex = 0;
            icsStateIndex = 0;
            cnsetSeenIndex = 0;
            cnsetSeenFAIIndex = 0;
            cnsetReadIndex = 0;
        }
    }
}