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
    /// <summary>
    /// A class used to store the response data of "CopyIntoItemsLocal" operation.
    /// </summary>
    public class CopyIntoItemsLocalResponse
    {   
        /// <summary>
        /// Represents the result status of "CopyIntoItemsLocal" operation.
        /// </summary>
        private uint copyIntoItemsLocalResultValue;

        /// <summary>
        /// Represents the results collection of copy actions.
        /// </summary>
        private CopyResult[] resultsCollection;

        /// <summary>
        /// Initializes a new instance of the CopyIntoItemsLocalResponse class.
        /// </summary>
        /// <param name="copyIntoItemsLocalResult">A parameter represents the result status of "CopyIntoItemsLocal" operation.</param>
        /// <param name="results">A parameter represents the copy results collection of the copy actions performed in "CopyIntoItemsLocal" operation.</param>
        public CopyIntoItemsLocalResponse(uint copyIntoItemsLocalResult, CopyResult[] results)
        {
            this.copyIntoItemsLocalResultValue = copyIntoItemsLocalResult;
            this.resultsCollection = results;
        }

        /// <summary>
        /// Gets the result status of the "CopyIntoItemsLocal" operation.
        /// </summary>
        public uint CopyIntoItemsLocalResult
        {
            get
            {
                return this.copyIntoItemsLocalResultValue;
            }
        }

        /// <summary>
        /// Gets the copy results collection.
        /// </summary>
        public CopyResult[] Results
        {
            get
            {
                return this.resultsCollection;
            }
        }
    }
}