namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using System.Security.Cryptography;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to process simple chunking.
    /// </summary>
    public class SimpleChunking : AbstractChunking
    {
        /// <summary>
        /// Initializes a new instance of the SimpleChunking class
        /// </summary>
        /// <param name="fileContent">The content of the file.</param>
        public SimpleChunking(byte[] fileContent)
            : base(fileContent)
        {
        }

        /// <summary>
        /// This method is used to chunk the file data.
        /// </summary>
        /// <returns>A list of LeafNodeObjectData.</returns>
        public override List<LeafNodeObject> Chunking()
        {
            int maxChunkSize = 1 * 1024 * 1024;
            List<LeafNodeObject> list = new List<LeafNodeObject>();
            LeafNodeObject.IntermediateNodeObjectBuilder builder = new LeafNodeObject.IntermediateNodeObjectBuilder();
            int chunkStart = 0;

            if (this.FileContent.Length <= maxChunkSize)
            {
                list.Add(builder.Build(this.FileContent, this.GetSignature(this.FileContent)));

                return list;
            }

            while (chunkStart < this.FileContent.Length)
            {
                int chunkLength = chunkStart + maxChunkSize >= this.FileContent.Length ? this.FileContent.Length - chunkStart : maxChunkSize;
                byte[] temp = AdapterHelper.GetBytes(this.FileContent, chunkStart, chunkLength);
                list.Add(builder.Build(temp, this.GetSignature(temp)));
                chunkStart += chunkLength;
            }

            return list;
        }

        /// <summary>
        /// This method is used to analyze the chunk for simple chunk.
        /// </summary>
        /// <param name="rootNode">Specify the root node object which is needed to be analyzed.</param>
        /// <param name="site">Specify the ITestSite instance.</param>
        public override void AnalyzeChunking(IntermediateNodeObject rootNode, ITestSite site)
        {
            if (rootNode.DataSize.DataSize <= 1024 * 1024)
            {
                // These would be something ignored for MOSS2010, when the file size is less than 1MB.
                // In this case the server will response data without signature, but to be consistent with the behavior,
                // The Simple chunk still will be used, but empty signature check will be ignored.
                return;
            }
            else if (rootNode.DataSize.DataSize <= 250 * 1024 * 1024)
            {
                foreach (LeafNodeObject interNode in rootNode.IntermediateNodeObjectList)
                {
                    SignatureObject expect = this.GetSignature(interNode.DataNodeObjectData.ObjectData);
                    SignatureObject realValue = interNode.Signature;
                    if (!expect.Equals(interNode.Signature))
                    {
                        site.Assert.Fail("Expect the signature value {0}, but actual value is {1}", expect.ToString(), realValue.ToString());
                    }

                    site.Assert.IsTrue(interNode.DataSize.DataSize <= 1024 * 1024, "The size of each chunk should be equal or less than 1MB for simple chunk.");

                    site.Assert.AreEqual<ulong>(
                            (ulong)interNode.GetContent().Count,
                            interNode.DataSize.DataSize,
                            "Expect the data size value equal the number represented by the chunk.");

                    site.Assert.IsNotNull(
                        interNode.DataNodeObjectData,
                        "The Object References array of the Intermediate Node Object associated with this Data Node Object MUST have a single entry which MUST be the Object ID of the Data Node Object.");
                }

                site.Log.Add(LogEntryKind.Debug, "All the intermediate signature value matches.");

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    MsfsshttpdCapture.VerifySimpleChunk(SharedContext.Current.Site);
                }
            }
            else
            {
                throw new NotImplementedException("When the file size is larger than 250MB, because the signature method is not implemented, the analysis method is also not implemented.");
            }
        }

        /// <summary>
        /// Get signature for the chunk.
        /// </summary>
        /// <param name="array">The data of the chunk.</param>
        /// <returns>The signature instance.</returns>
        private SignatureObject GetSignature(byte[] array)
        {
            if (this.FileContent.Length <= 250 * 1024 * 1024)
            {
                SHA1 sha = new SHA1CryptoServiceProvider();
                byte[] temp = sha.ComputeHash(array);
                sha.Dispose();

                SignatureObject signature = new SignatureObject();
                signature.SignatureData = new BinaryItem(temp);
                return signature;
            }
            else
            {
                throw new NotImplementedException("When the file size is larger than 250MB, the signature method is not implemented.");
            }
        }
    }
}