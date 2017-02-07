namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Cryptography;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to process zip file chunking.
    /// </summary>
    public class ZipFilesChunking : AbstractChunking
    {
        /// <summary>
        /// Initializes a new instance of the ZipFilesChunking class
        /// </summary>
        /// <param name="fileContent">The content of the file.</param>
        public ZipFilesChunking(byte[] fileContent)
            : base(fileContent)
        {
        }

        /// <summary>
        /// This method is used to chunk the file data.
        /// </summary>
        /// <returns>A list of LeafNodeObjectData.</returns>
        public override List<LeafNodeObject> Chunking()
        {
            List<LeafNodeObject> list = new List<LeafNodeObject>();
            LeafNodeObject.IntermediateNodeObjectBuilder builder = new LeafNodeObject.IntermediateNodeObjectBuilder();

            int index = 0;
            while (ZipHeader.IsFileHeader(this.FileContent, index))
            {
                byte[] dataFileSignatureBytes;
                byte[] header = this.AnalyzeFileHeader(this.FileContent, index, out dataFileSignatureBytes);
                int headerLength = header.Length;
                int compressedSize = (int)this.GetCompressedSize(dataFileSignatureBytes);

                if (headerLength + compressedSize <= 4096)
                {
                    list.Add(builder.Build(AdapterHelper.GetBytes(this.FileContent, index, headerLength + compressedSize), this.GetSingleChunkSignature(header, dataFileSignatureBytes)));
                    index += headerLength += compressedSize;
                }
                else
                {
                    list.Add(builder.Build(header, this.GetSHA1Signature(header)));
                    index += headerLength;

                    byte[] dataFile = AdapterHelper.GetBytes(this.FileContent, index, compressedSize);

                    if (dataFile.Length <= 1048576)
                    {
                        list.Add(builder.Build(dataFile, this.GetDataFileSignature(dataFileSignatureBytes)));
                    }
                    else
                    {
                        list.AddRange(this.GetSubChunkList(dataFile));
                    }

                    index += compressedSize;
                }
            }

            if (0 == index)
            {
                return null;
            }

            byte[] final = AdapterHelper.GetBytes(this.FileContent, index, this.FileContent.Length - index);

            if (final.Length <= 1048576)
            {
                list.Add(builder.Build(final, this.GetSHA1Signature(final)));
            }
            else
            {
                // In current, it has no idea about how to compute the signature for final part larger than 1MB.
                throw new NotImplementedException("If the final chunk is larger than 1MB, the signature method is not implemented.");
            }

            return list;
        }

        /// <summary>
        /// This method is used to analyze the chunk.
        /// </summary>
        /// <param name="rootNode">Specify the root node object which will be analyzed.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public override void AnalyzeChunking(IntermediateNodeObject rootNode, ITestSite site)
        {
            List<LeafNodeObject> cloneList = new List<LeafNodeObject>(rootNode.IntermediateNodeObjectList);

            while (cloneList.Count != 0)
            {
                LeafNodeObject nodeObject = cloneList.First();
                byte[] content = nodeObject.DataNodeObjectData.ObjectData;

                if (cloneList.Count == 1)
                {
                    if (content.Length > 1048576)
                    {
                        throw new NotImplementedException("If the final chunk is larger than 1MB, the signature method is not implemented.");
                    }

                    // Only final chunk left
                    SignatureObject expect = this.GetSHA1Signature(content);
                    if (!expect.Equals(nodeObject.Signature))
                    {
                        site.Assert.Fail("For the Zip file, final part chunk expect the signature {0}, actual signature {1}", expect.ToString(), nodeObject.Signature.ToString());
                    }

                    // Verify the less than 1MB final part related requirements
                    if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                    {
                        MsfsshttpdCapture.VerifySmallFinalChunk(SharedContext.Current.Site);
                    }
                }
                else
                {
                    if (ZipHeader.IsFileHeader(content, 0))
                    {
                        byte[] dataFileSignatureBytes;
                        byte[] header = this.AnalyzeFileHeader(content, 0, out dataFileSignatureBytes);
                        int headerLength = header.Length;
                        int compressedSize = (int)this.GetCompressedSize(dataFileSignatureBytes);

                        if (headerLength + compressedSize <= 4096)
                        {
                            if (Common.GetConfigurationPropertyValue("SutVersion", SharedContext.Current.Site) != "SharePointFoundation2010" && Common.GetConfigurationPropertyValue("SutVersion", SharedContext.Current.Site) != "SharePointServer2010")
                            {
                                LeafNodeObject expectNode = new LeafNodeObject.IntermediateNodeObjectBuilder().Build(content, this.GetSingleChunkSignature(header, dataFileSignatureBytes));
                                if (!expectNode.Signature.Equals(nodeObject.Signature))
                                {
                                    site.Assert.Fail("For the Zip file, when zip file is less than 4096, expect the signature {0}, actual signature {1}", expectNode.Signature.ToString(), nodeObject.Signature.ToString());
                                }
                            }

                            // Verify the zip file less than 4096 bytes
                            MsfsshttpdCapture.VerifyZipFileLessThan4096Bytes(SharedContext.Current.Site);
                        }
                        else
                        {
                            SignatureObject expectHeader = this.GetSHA1Signature(header);
                            if (!expectHeader.Equals(nodeObject.Signature))
                            {
                                site.Assert.Fail("For the Zip file header, expect the signature {0}, actual signature {1}", expectHeader.ToString(), nodeObject.Signature.ToString());
                            }

                            // Remove the header node
                            cloneList.RemoveAt(0);

                            // Then expect the next is file content node
                            nodeObject = cloneList.First();

                            // Here having something need to be distinguished between the MOSS2010 and MOSS2013
                            if (nodeObject.DataNodeObjectData == null && nodeObject.IntermediateNodeObjectList != null)
                            {
                                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8213, SharedContext.Current.Site))
                                {
                                    bool isR8213Verified = false;
                                    if (compressedSize > 1024 * 1024 && nodeObject.IntermediateNodeObjectList.Count > 1)
                                    {
                                        isR8213Verified = true;
                                    }

                                    site.CaptureRequirementIfIsTrue(
                                            isR8213Verified,
                                            "MS-FSSHTTPD",
                                            8213,
                                            @"[In Appendix A: Product Behavior] For implementation, if the number of .ZIP file bytes represented by a chunk is greater than 1 megabyte, a list of subchunks is generated. <4> Section 2.4.1:  For SharePoint Server 2010, if the number of .ZIP file bytes represented by a chunk is greater than 1 megabyte, a list of subchunks is generated.");
                                }

                                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8207, SharedContext.Current.Site))
                                {
                                    bool isR8207Verified = false;
                                    if (compressedSize > 3 * 1024 * 1024 && nodeObject.IntermediateNodeObjectList.Count > 1)
                                    {
                                        isR8207Verified = true;
                                    }

                                    site.CaptureRequirementIfIsTrue(
                                            isR8207Verified,
                                            "MS-FSSHTTPD",
                                            8207,
                                            @"[In Appendix A: Product Behavior] For implementation, if the number of .ZIP file bytes represented by a chunk is greater than 3 megabytes, a list of subchunks is generated. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010/Microsft Office 2016/Microsft SharePoint Server 2016 follow this behavior.)");
                                }

                                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8208, SharedContext.Current.Site) && Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8210, SharedContext.Current.Site))
                                {
                                    bool isR8208Verified = true;
                                    bool isR8210Verified = true;
                                    if (nodeObject.IntermediateNodeObjectList[nodeObject.IntermediateNodeObjectList.Count - 1].DataSize.DataSize > 1024 * 1024)
                                    {
                                        isR8208Verified = false;
                                    }

                                    for (int i = 0; i < nodeObject.IntermediateNodeObjectList.Count - 1; i++)
                                    {
                                        if (nodeObject.IntermediateNodeObjectList[i].DataSize.DataSize != 1024 * 1024)
                                        {
                                            isR8210Verified = false;
                                        }
                                    }

                                    site.CaptureRequirementIfIsTrue(
                                            isR8208Verified,
                                            "MS-FSSHTTPD",
                                            8208,
                                            @"[In Appendix A: Product Behavior] The size of each subchunk is at most 1 megabyte. (Microsoft SharePoint Server 2010 follows this behavior.)");

                                    site.CaptureRequirementIfIsTrue(
                                            isR8210Verified,
                                            "MS-FSSHTTPD",
                                            8210,
                                            @"[In Appendix A: Product Behavior] All but the last subchunk MUST be 1 megabyte in size. (Microsfot SharePoint Server 2010 follows this behavior.)");
                                }

                                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8209, SharedContext.Current.Site) && Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 8211, SharedContext.Current.Site))
                                {
                                    bool isR8209Verified = true;
                                    bool isR8211Verified = true;
                                    if (nodeObject.IntermediateNodeObjectList[nodeObject.IntermediateNodeObjectList.Count - 1].DataSize.DataSize > 3 * 1024 * 1024)
                                    {
                                        isR8209Verified = false;
                                    }

                                    for (int i = 0; i < nodeObject.IntermediateNodeObjectList.Count - 1; i++)
                                    {
                                        if (nodeObject.IntermediateNodeObjectList[i].DataSize.DataSize != 3 * 1024 * 1024)
                                        {
                                            isR8211Verified = false;
                                        }
                                    }

                                    site.CaptureRequirementIfIsTrue(
                                            isR8209Verified,
                                            "MS-FSSHTTPD",
                                            8209,
                                            @"[In Appendix A: Product Behavior] The size of each subchunk is at most 3 megabytes. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010/Microsft Office 2016/Microsft SharePoint Server 2016 follow this behavior.)");

                                    site.CaptureRequirementIfIsTrue(
                                            isR8211Verified,
                                            "MS-FSSHTTPD",
                                            8211,
                                            @"[In Appendix A: Product Behavior] All but the last subchunk MUST be 3 megabyte in size. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010/Microsft Office 2016/Microsft SharePoint Server 2016 follow this behavior.)");
                                }
                            }
                            else if (nodeObject.DataNodeObjectData != null)
                            {
                                site.Assert.AreEqual<ulong>(
                                            (ulong)compressedSize,
                                            nodeObject.DataSize.DataSize,
                                            "The Data Size of the Intermediate Node Object MUST be the total number of bytes represented by the chunk.");

                                if (Common.GetConfigurationPropertyValue("SutVersion", SharedContext.Current.Site) != "SharePointFoundation2010" && Common.GetConfigurationPropertyValue("SutVersion", SharedContext.Current.Site) != "SharePointServer2010")
                                {
                                    SignatureObject contentSignature = new SignatureObject();
                                    contentSignature.SignatureData = new BinaryItem(dataFileSignatureBytes);
                                    if (!contentSignature.Equals(nodeObject.Signature))
                                    {
                                        site.Assert.Fail("For the Zip file content, expect the signature {0}, actual signature {1}", contentSignature.ToString(), nodeObject.Signature.ToString());
                                    }
                                }

                                // Verify the zip file larger than 4096 bytes and less than 1MB.
                                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                                {
                                    MsfsshttpdCapture.VerifyZipFileHeaderAndContentSignature(SharedContext.Current.Site);
                                    MsfsshttpdCapture.VerifyIntermediateNodeForZipFileChunk(SharedContext.Current.Site);
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException("The DataNodeObjectData and IntermediateNodeObjectList cannot be null at the same time.");
                            }
                        }
                    }
                }

                cloneList.RemoveAt(0);
            }
        }

        /// <summary>
        /// Convert chunk data to LeafNodeObjectData from byte array.
        /// </summary>
        /// <param name="chunkData">A byte array that contains the data.</param>
        /// <returns>A list of LeafNodeObjectData.</returns>
        private List<LeafNodeObject> GetSubChunkList(byte[] chunkData)
        {
            List<LeafNodeObject> subChunkList = new List<LeafNodeObject>();
            int index = 0;
            while (index < chunkData.Length)
            {
                int length = chunkData.Length - index < 1048576 ? chunkData.Length - index : 1048576;
                byte[] temp = AdapterHelper.GetBytes(chunkData, index, length);
                subChunkList.Add(new LeafNodeObject.IntermediateNodeObjectBuilder().Build(temp, this.GetSubChunkSignature()));
                index += length;
            }

            return subChunkList;
        }

        /// <summary>
        /// This method is used to analyze the zip file header.
        /// </summary>
        /// <param name="content">Specify the zip content.</param>
        /// <param name="index">Specify the start position.</param>
        /// <param name="dataFileSignature">Specify the output value for the data file signature.</param>
        /// <returns>Return the data file content.</returns>
        private byte[] AnalyzeFileHeader(byte[] content, int index, out byte[] dataFileSignature)
        {
            int crc32 = BitConverter.ToInt32(content, index + 14);
            int compressedSize = BitConverter.ToInt32(content, index + 18);
            int uncompressedSize = BitConverter.ToInt32(content, index + 22);
            int fileNameLength = BitConverter.ToInt16(content, index + 26);
            int extraFileldLength = BitConverter.ToInt16(content, index + 28);
            int headerLength = 30 + fileNameLength + extraFileldLength;

            BitWriter writer = new BitWriter(20);
            writer.AppendInit32(crc32, 32);
            writer.AppendUInt64((ulong)compressedSize, 64);
            writer.AppendUInt64((ulong)uncompressedSize, 64);
            dataFileSignature = writer.Bytes;

            return AdapterHelper.GetBytes(content, index, headerLength);
        }

        /// <summary>
        /// This method is used to get the compressed size value from the data file signature.
        /// </summary>
        /// <param name="dataFileSignature">Specify the signature of the zip file content.</param>
        /// <returns>Return the compressed size value.</returns>
        private ulong GetCompressedSize(byte[] dataFileSignature)
        {
            using (BitReader reader = new BitReader(dataFileSignature, 0))
            {
                reader.ReadUInt32(32);
                return reader.ReadUInt64(64);
            }
        }

        /// <summary>
        /// Get the signature for single chunk.
        /// </summary>
        /// <param name="header">The data of file header.</param>
        /// <param name="dataFile">The data of data file.</param>
        /// <returns>An instance of SignatureObject.</returns>
        private SignatureObject GetSingleChunkSignature(byte[] header, byte[] dataFile)
        {
            SHA1 sha = new SHA1CryptoServiceProvider();
            byte[] headerSignature = sha.ComputeHash(header);
            sha.Dispose();
            byte[] singleSignature = null;

            if (SharedContext.Current.CellStorageVersionType.MinorVersion >= 2)
            {
                singleSignature = new byte[dataFile.Length];

                for (int i = 0; i < headerSignature.Length; i++)
                {
                    singleSignature[i] = (byte)(headerSignature[i] ^ dataFile[i]);
                }
            }
            else
            {
                List<byte> tmp = new List<byte>();
                tmp.AddRange(headerSignature);
                tmp.AddRange(dataFile);

                singleSignature = tmp.ToArray(); 
            }

            SignatureObject signature = new SignatureObject();
            signature.SignatureData = new BinaryItem(singleSignature);

            return signature;
        }

        /// <summary>
        /// Get signature with SHA1 algorithm.
        /// </summary>
        /// <param name="array">The input data.</param>
        /// <returns>An instance of SignatureObject.</returns>
        private SignatureObject GetSHA1Signature(byte[] array)
        {
            SHA1 sha = new SHA1CryptoServiceProvider();
            byte[] temp = sha.ComputeHash(array);
            sha.Dispose();

            SignatureObject signature = new SignatureObject();
            signature.SignatureData = new BinaryItem(temp);
            return signature;
        }

        /// <summary>
        /// Get the signature for data file.
        /// </summary>
        /// <param name="array">The input data.</param>
        /// <returns>An instance of SignatureObject.</returns>
        private SignatureObject GetDataFileSignature(byte[] array)
        {
            SignatureObject signature = new SignatureObject();
            signature.SignatureData = new BinaryItem(array);

            return signature;
        }

        /// <summary>
        /// Get the signature for sub chunk.
        /// </summary>
        /// <returns>An instance of SignatureObject.</returns>
        private SignatureObject GetSubChunkSignature()
        {
            // In current, it has no idea about how to compute the signature for sub chunk.
            throw new NotImplementedException("The Get sub chunk signature method is not implemented.");
        }
    }
}