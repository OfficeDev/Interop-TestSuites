//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class is used to create instance of AbstractChunking.
    /// </summary>
    public class ChunkingFactory
    {
        /// <summary>
        /// Prevents a default instance of the ChunkingFactory class from being created
        /// </summary>
        private ChunkingFactory()
        { 
        }

        /// <summary>
        /// This method is used to create the instance of AbstractChunking.
        /// </summary>
        /// <param name="fileContent">The content of the file.</param>
        /// <returns>The instance of AbstractChunking.</returns>
        public static AbstractChunking CreateChunkingInstance(byte[] fileContent)
        {
            if (ZipHeader.IsFileHeader(fileContent, 0))
            {
                return new ZipFilesChunking(fileContent);
            }
            else
            {
                return new RDCAnalysisChunking(fileContent);
            }
        }

        /// <summary>
        /// This method is used to create the instance of AbstractChunking.
        /// </summary>
        /// <param name="nodeObject">Specify the root node object.</param>
        /// <returns>The instance of AbstractChunking.</returns>
        public static AbstractChunking CreateChunkingInstance(RootNodeObject nodeObject)
        {
            byte[] fileContent = nodeObject.GetContent().ToArray();

            if (EditorsTableUtils.IsEditorsTableHeader(fileContent))
            {
                return null;
            }

            if (ZipHeader.IsFileHeader(fileContent, 0))
            {
                return new ZipFilesChunking(fileContent);
            }
            else
            {
                // For SharePoint Server 2013 compatible SUTs, always using the RDC Chunking method in the current test suite involved file resources.
                if (SharedContext.Current.CellStorageVersionType.MinorVersion >= 2)
                {
                    return new RDCAnalysisChunking(fileContent);
                }

                // For SharePoint Server 2010 SP2 compatible SUTs, chunking method depends on file content and size. So first try using the simple chunking.  
                AbstractChunking returnChunking = new SimpleChunking(fileContent);

                List<IntermediateNodeObject> nodes = returnChunking.Chunking();
                if (nodeObject.IntermediateNodeObjectList.Count == nodes.Count)
                {
                    bool isDataSizeMatching = true;
                    for (int i = 0; i < nodes.Count; i++)
                    {
                        if (nodeObject.IntermediateNodeObjectList[i].DataSize.DataSize != nodes[i].DataSize.DataSize)
                        {
                            isDataSizeMatching = false;
                            break;
                        }
                    }

                    if (isDataSizeMatching)
                    {
                        return returnChunking;
                    }
                }

                // If the intermediate count number or data size does not equals, then try to use RDC chunking method.
                return new RDCAnalysisChunking(fileContent);
            }
        }

        /// <summary>
        /// This method is used to create the instance of AbstractChunking.
        /// </summary>
        /// <param name="fileContent">The content of the file.</param>
        /// <param name="chunkingMethod">The type of chunking methods.</param>
        /// <returns>The instance of AbstractChunking.</returns>
        public static AbstractChunking CreateChunkingInstance(byte[] fileContent, ChunkingMethod chunkingMethod)
        {
            AbstractChunking chunking = null;
            switch (chunkingMethod)
            {
                case ChunkingMethod.RDCAnalysis:
                    chunking = new RDCAnalysisChunking(fileContent);
                    break;
                case ChunkingMethod.SimpleAlgorithm:
                    chunking = new SimpleChunking(fileContent);
                    break;
                case ChunkingMethod.ZipAlgorithm:
                    chunking = new ZipFilesChunking(fileContent);
                    break;

                default:
                    throw new InvalidOperationException("Cannot support the chunking type" + chunkingMethod.ToString());
            }

            return chunking;
        }
    }

    /// <summary>
    /// This class is used to check is this a zip file header.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public sealed class ZipHeader
    {
        /// <summary>
        /// The file header in zip.
        /// </summary>
        public static readonly byte[] LocalFileHeader = new byte[] { 0x50, 0x4b, 0x03, 0x04 };

        /// <summary>
        /// Prevents a default instance of the ZipHeader class from being created
        /// </summary>
        private ZipHeader()
        { 
        }

        /// <summary>
        /// Check the input data is a local file header.
        /// </summary>
        /// <param name="byteArray">The content of a file.</param>
        /// <param name="index">The index where to start.</param>
        /// <returns>True if the input data is a local file header, otherwise false.</returns>
        public static bool IsFileHeader(byte[] byteArray, int index)
        {
            if (AdapterHelper.ByteArrayEquals(LocalFileHeader, AdapterHelper.GetBytes(byteArray, index, 4)))
            {
                return true;
            }

            return false;
        }
    }
}