namespace Microsoft.Protocols.TestSuites.MS_OXWSITEMID
{
    using System;
    using System.IO;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// Adapter class of MS_OXWSITEMID.
    /// </summary>
    public class MS_OXWSITEMIDAdapter : ManagedAdapterBase, IMS_OXWSITEMIDAdapter
    {
        #region Initialize TestSuite
        /// <summary>
        /// Initialize some variables overridden.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite Class.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
        }

        #endregion

        /// <summary>
        /// Parse an ItemId's Id from a base64 string to a ItemIdId object according to the defined format
        /// </summary>
        /// <param name="itemId">An ItemIdType object</param>
        /// <returns>An ItemIdId object as the result of parsing</returns>
        public ItemIdId ParseItemId(ItemIdType itemId)
        {
            Site.Assert.IsNotNull(itemId, "The input itemId should not be null.");
            ItemIdId parsedId = new ItemIdId();
            byte[] id = Convert.FromBase64String(itemId.Id);
            int currentIndex = 0;

            // [Byte] Compression Type
            parsedId.CompressionByte = id[currentIndex];
            if (parsedId.CompressionByte == 1)
            {
                // The max length of Base64-encoded Id is 512, which is 512*3/4=384 after decoding
                // Then minus 1 (Compress Byte),which is 383
                id = this.Decompress(id, 383).ToArray();
            }
            else
            {
                currentIndex++;
            }

            // [Byte] Id Storage Type
            Site.Assert.IsTrue(
                Enum.IsDefined(typeof(IdStorageType), id[currentIndex]),
               "The id storage type should be valid. Actually the value is '{0}'.",
               id[currentIndex]);
            parsedId.StorageType = (IdStorageType)id[currentIndex];
            currentIndex++;
            switch (parsedId.StorageType)
            {
                case IdStorageType.MailboxItemSmtpAddressBased:
                case IdStorageType.MailboxItemMailboxGuidBased:
                case IdStorageType.ConversationIdMailboxGuidBased:
                    {
                        parsedId.MonikerLength = BitConverter.ToInt16(id, currentIndex);
                        currentIndex += 2;

                        parsedId.MonikerBytes = new byte[(int)parsedId.MonikerLength];
                        Array.Copy(id, currentIndex, parsedId.MonikerBytes, 0, (int)parsedId.MonikerLength);
                        currentIndex += (int)parsedId.MonikerLength;

                        if (Enum.IsDefined(typeof(IdProcessingInstructionType), id[currentIndex]))
                        {
                            parsedId.IdProcessingInstruction = (IdProcessingInstructionType)id[currentIndex];
                            currentIndex++;
                        }
                        else 
                        {
                            Site.Assert.Fail("Undefined Id Processing Instruction Type value {0}", id[currentIndex]);
                        }

                        parsedId.StoreIdLength = BitConverter.ToInt16(id, currentIndex);
                        currentIndex += 2;
                        parsedId.StoreId = new byte[parsedId.StoreIdLength];
                        Array.Copy(id, currentIndex, parsedId.StoreId, 0, parsedId.StoreIdLength);
                        currentIndex += parsedId.StoreIdLength;

                        break;
                    }

                case IdStorageType.PublicFolder:
                case IdStorageType.ActiveDirectoryObject:
                    {
                        parsedId.StoreIdLength = BitConverter.ToInt16(id, currentIndex);
                        currentIndex += 2;

                        parsedId.StoreId = new byte[parsedId.StoreIdLength];
                        Array.Copy(id, currentIndex, parsedId.StoreId, 0, parsedId.StoreIdLength);
                        currentIndex += parsedId.StoreIdLength;

                        break;
                    }

                case IdStorageType.PublicFolderItem:
                    {
                        if (Enum.IsDefined(typeof(IdProcessingInstructionType), id[currentIndex]))
                        {
                            parsedId.IdProcessingInstruction = (IdProcessingInstructionType)id[currentIndex];
                            currentIndex++;
                        }
                        else
                        {
                            Site.Assert.Fail("Undefined Id Processing Instruction Type value {0}", id[currentIndex]);
                        }

                        parsedId.StoreIdLength = BitConverter.ToInt16(id, currentIndex);
                        currentIndex += 2;

                        parsedId.StoreId = new byte[parsedId.StoreIdLength];
                        Array.Copy(id, currentIndex, parsedId.StoreId, 0, parsedId.StoreIdLength);
                        currentIndex += parsedId.StoreIdLength;

                        parsedId.FolderIdLength = BitConverter.ToInt16(id, currentIndex);
                        currentIndex += 2;

                        parsedId.FolderId = new byte[(int)parsedId.FolderIdLength];
                        Array.Copy(id, currentIndex, parsedId.FolderId, 0, (int)parsedId.FolderIdLength);
                        currentIndex += (int)parsedId.FolderIdLength;

                        break;
                    }

                default:
                    {
                        Site.Assert.Fail("No format defined for Id Storage Type {0}", parsedId.StorageType.ToString("g"));
                        break;
                    }
            }

            if (currentIndex < id.Length)
            {
                parsedId.AttachmentIdCount = id[currentIndex];
                currentIndex++;

                short attachmentIdCount = Convert.ToInt16(parsedId.AttachmentIdCount);
                parsedId.AttachmentIds = new AttachmentId[attachmentIdCount];

                for (int i = 0; i < attachmentIdCount; i++)
                {
                    parsedId.AttachmentIds[i].AttachmentIdLength = BitConverter.ToInt16(id, currentIndex);
                    currentIndex += 2;

                    parsedId.AttachmentIds[i].Id = new byte[parsedId.AttachmentIds[i].AttachmentIdLength];
                    Array.Copy(id, currentIndex, parsedId.AttachmentIds[i].Id, 0, parsedId.AttachmentIds[i].AttachmentIdLength);
                    currentIndex += parsedId.AttachmentIds[i].AttachmentIdLength;
                }
            }

            Site.Assert.AreEqual<int>(id.Length, currentIndex, "There should be no bytes left after parsing item id!");
            return parsedId;
        }

        /// <summary>
        /// Simple RLE compressor for item IDs. Bytes that do not repeat are written directly.
        /// Bytes that repeat more than once are written twice, followed by the number of 
        /// additional times to write the byte (i.e., total run length minus two).
        /// </summary>
        /// <param name="streamIn">input stream to compress</param>
        /// <param name="compressorId">id of the compressor</param>
        /// <returns>compressed bytes</returns>
        public byte[] Compress(byte[] streamIn, byte compressorId)
        {
            byte[] streamOut = new byte[streamIn.Length];
            int index = 0;
            streamOut[index++] = compressorId;
            if (index == streamIn.Length)
            {
                return streamIn;
            }

            // Ignore the first byte, because it is a placeholder for the compression tag.
            // Keep a placeholder so that, if the caller ends up not doing any compression
            // at all, they can simply put the compression tag for "NoCompression" in the 
            // first byte and everything works.
            byte[] input = streamIn;

            for (int runStart = 1; runStart < (int)streamIn.Length; /* runStart incremented below */)
            {
                // Always write the start character.
                streamOut[index++] = input[runStart];
                if (index == streamIn.Length)
                {
                    return streamIn;
                }

                // Now look for a run of more than one character. The maximum run to be 
                // handled at once is the maximum value that can be written out in an 
                // (unsigned) byte _or_ the maximum remaining input, whichever is smaller.
                // One caveat is that only the run length _minus two_ is written, 
                // because the two characters that indicate a run are not written. So 
                // Byte.MaxValue + 2 can be handled.
                int maxRun = Math.Min(byte.MaxValue + 2, (int)streamIn.Length - runStart);
                int runLength = 1;
                for (runLength = 1;
                    runLength < maxRun && input[runStart] == input[runStart + runLength];
                    ++runLength)
                {
                    // Nothing.
                }

                // Is this a run of more than one byte?
                if (runLength > 1)
                {
                    // Yes, write the byte again, followed by the number of additional
                    // times to write the byte (which is the total run length minus 2,
                    // because the byte has already been written twice).
                    streamOut[index++] = input[runStart];
                    if (index == streamIn.Length)
                    {
                        return streamIn;
                    }

                    if (runLength > byte.MaxValue + 2)
                    {
                        Site.Assert.Fail("Total run length exceeds. The max number of continuous same bytes is byte.MaxValue+2, but actually it is {0}", runLength);
                    }

                    streamOut[index++] = (byte)(runLength - 2);
                    if (index == streamIn.Length)
                    {
                        return streamIn;
                    }
                }

                // Move to the first byte following the run.
                runStart += runLength;
            }
  
            byte[] outBytes = new byte[index];
            Array.Copy(streamOut, outBytes, index);
            return outBytes;
        }

        /// <summary>
        /// Decompresses the passed byte array using RLE scheme.
        /// </summary>
        /// <param name="input">Bytes to decompress</param>
        /// <param name="maxLength">Max allowed length for the byte array</param>
        /// <returns>Decompressed bytes minus the first byte of input</returns>
        public MemoryStream Decompress(byte[] input, int maxLength)
        {
            // It can't be assumed that the compressed data size must be less than maxLength.
            // If the compressed data consists of a series of double characters
            // followed by a 0 character count, compressed data will be larger than 
            // decompressed. (i.e. xx0 decompresses to xx.)
            int initialStreamSize = Math.Min(input.Length, maxLength);

            MemoryStream stream = new MemoryStream(initialStreamSize);
            BinaryWriter writer = new BinaryWriter(stream);

            // Ignore the first byte, which the caller used to identify the compression
            // scheme.
            for (int i = 1; i < input.Length; ++i)
            {
                // If this byte differs from the following one (or it's at the end of the
                //  array), then just output the byte.
                if (i == input.Length - 1 ||
                    input[i] != input[i + 1])
                {
                    writer.Write(input[i]);
                }
                else
                {
                    // Because repeat characters are always followed by a character count,
                    // if i == input.Length - 2, the character count is missing & the id is 
                    // invalid.
                    if (i == input.Length - 2)
                    {
                        Site.Assert.Fail("Invalid Id format. Some bytes are missing at the end of compressed Id");
                    }

                    // The bytes are the same. Read the third byte to see how many additional
                    // times to write the byte (over and above the two that are already 
                    // there).
                    byte runLength = input[i + 2];
                    for (int j = 0; j < runLength + 2; ++j)
                    {
                        writer.Write(input[i]);
                    }

                    // Skip the duplicate byte and the run length.
                    i += 2;
                }

                if (stream.Length > maxLength)
                {
                    Site.Assert.Fail("Invalid Id format. The actually Id length {0} exceeds the max value {1}", maxLength);
                }
            }

            writer.Flush();
            stream.Position = 0L;

            return stream;
        }
    }
}