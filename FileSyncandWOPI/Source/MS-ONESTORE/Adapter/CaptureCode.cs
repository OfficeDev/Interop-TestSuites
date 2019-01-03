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
    using System.IO;
    using System.Collections.Generic;
    using TestTools;
    using System.Linq;

    /// <summary>
    /// The class provides the methods to write capture code.
    /// </summary>
    public partial class MS_ONESTOREAdapter
    {
        /// <summary>
        /// This method is used to verify the requirements related with MS-ONESTORE.
        /// </summary>
        /// <param name="file">The instance of revision-based format file.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        public void VerifyRevisionStoreFile(OneNoteRevisionStoreFile file, string fileName, ITestSite site)
        {
            this.VerifyHeader(file, fileName, site);
            this.VerifyFreeChunkList(file.FreeChunkList, site);
            this.VerifyTransactionLog(file, site);
            this.VerifyHashedChunkList(file.HashedChunkList, site);
            this.VerifyRootFileNodeList(file.RootFileNodeList, fileName, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with Header structure.
        /// </summary>
        /// <param name="header">The instacne of header structure.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyHeader(OneNoteRevisionStoreFile file, string fileName, ITestSite site)
        {
            Header header = file.Header;
            string extension = Path.GetExtension(fileName);

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R217
            site.CaptureRequirement(
                     217,
                     @"[In Header] The Header structure MUST be at the beginning of the file.");

            #region Verify the guidFileType
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R219
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidFileType,
                    typeof(Guid),
                    219,
                    @"[In Header] guidFileType (16 bytes): A GUID, as specified by [MS-DTYP], that specifies the type of the revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R220
            site.CaptureRequirementIfIsTrue(
                    header.guidFileType == Guid.Parse("{7B5C52E4-D88C-4DA7-AEB1-5378D02996D3}") ||
                    header.guidFileType == Guid.Parse("{43FF2FA1-EFD9-4C76-9EE2-10EA5722765F}"),
                    220,
                    @"[In Header] [guidFileType] MUST be one of the values from the following table[ {7B5C52E4-D88C-4DA7-AEB1-5378D02996D3} and {43FF2FA1-EFD9-4C76-9EE2-10EA5722765F}].");
            #endregion

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R223
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidFile,
                    typeof(Guid),
                    223,
                    @"[In Header] guidFile (16 bytes): A GUID, as specified by [MS-DTYP], that specifies the identity of this revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R225
            site.CaptureRequirementIfAreEqual<Guid>(
                    Guid.Parse("{00000000-0000-0000-0000-000000000000}"),
                    header.guidLegacyFileVersion,
                    225,
                    @"[In Header] guidLegacyFileVersion (16 bytes): MUST be ""{00000000-0000-0000-0000-000000000000}"" and MUST be ignored.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R226
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidFileFormat,
                    typeof(Guid),
                    226,
                    @"[In Header] guidFileFormat (16 bytes): A GUID, as specified by [MS-DTYP], that specifies that the file is a revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R227
            site.CaptureRequirementIfAreEqual<Guid>(
                    Guid.Parse("{109ADD3F-911B-49F5-A5D0-1791EDC8AED8}"),
                    header.guidFileFormat,
                    227,
                    @"[In Header] [guidFileFormat]: MUST be ""{109ADD3F-911B-49F5-A5D0-1791EDC8AED8}"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R228
            site.CaptureRequirementIfIsInstanceOfType(
                    header.ffvLastCodeThatWroteToThisFile,
                    typeof(uint),
                    228,
                    @"[In Header] ffvLastCodeThatWroteToThisFile (4 bytes): An unsigned integer.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R229
            site.CaptureRequirementIfIsTrue(
                    header.ffvLastCodeThatWroteToThisFile == 0x0000002A || header.ffvLastCodeThatWroteToThisFile == 0x0000001B,
                    229,
                    @"[In Header] [ffvLastCodeThatWroteToThisFile] MUST be one of the values in the following table[0x0000002A and 0x0000001B], depending on the file type.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R232
            site.CaptureRequirementIfIsInstanceOfType(
                    header.ffvOldestCodeThatHasWrittenToThisFile,
                    typeof(uint),
                    232,
                    @"[In Header] ffvOldestCodeThatHasWrittenToThisFile (4 bytes): An unsigned integer.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R233
            site.CaptureRequirementIfIsTrue(
                    header.ffvOldestCodeThatHasWrittenToThisFile == 0x0000002A || header.ffvOldestCodeThatHasWrittenToThisFile == 0x0000001B,
                    233,
                    @"[In Header] [ffvOldestCodeThatHasWrittenToThisFile] MUST be one of the values in the following table[0x0000002A and 0x0000001B], depending on the file format of this file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R236
            site.CaptureRequirementIfIsInstanceOfType(
                    header.ffvNewestCodeThatHasWrittenToThisFile,
                    typeof(uint),
                    236,
                    @"[In Header] ffvNewestCodeThatHasWrittenToThisFile (4 bytes): An unsigned integer.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R237
            site.CaptureRequirementIfIsTrue(
                    header.ffvNewestCodeThatHasWrittenToThisFile == 0x0000002A || header.ffvNewestCodeThatHasWrittenToThisFile == 0x0000001B,
                    237,
                    @"[In Header]  [ffvNewestCodeThatHasWrittenToThisFile] MUST be one of the values in the following table[0x0000002A and 0x0000001B], depending on the file format of this file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R240
            site.CaptureRequirementIfIsInstanceOfType(
                    header.ffvOldestCodeThatMayReadThisFile,
                    typeof(uint),
                    240,
                    @"[In Header] ffvOldestCodeThatMayReadThisFile (4 bytes): An unsigned integer.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R241
            site.CaptureRequirementIfIsTrue(
                    header.ffvOldestCodeThatMayReadThisFile == 0x0000002A || header.ffvOldestCodeThatMayReadThisFile == 0x0000001B,
                    241,
                    @"[In Header] [ffvOldestCodeThatMayReadThisFile] MUST be one of the values in the following table[0x0000002A and 0x0000001B], depending on the file format of this file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R244
            site.CaptureRequirementIfIsTrue(
                    header.fcrLegacyFreeChunkList.GetType() == typeof(FileChunkReference32) && header.fcrLegacyFreeChunkList.IsfcrZero(),
                    244,
                    @"[In Header] fcrLegacyFreeChunkList (8 bytes): A FileChunkReference32 structure (section 2.2.4.1) that MUST have a value of ""fcrZero"" (see section 2.2.4).");

            this.VerifyFileChunkReference32(header.fcrLegacyFreeChunkList, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R245
            site.CaptureRequirementIfIsTrue(
                    header.fcrLegacyTransactionLog.GetType() == typeof(FileChunkReference32) && header.fcrLegacyTransactionLog.IsfcrNil(),
                    245,
                    @"[In Header] fcrLegacyTransactionLog (8 bytes): A FileChunkReference32 structure that MUST be ""fcrNil"" (see section 2.2.4).");

            this.VerifyFileChunkReference32(header.fcrLegacyTransactionLog, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R246
            site.CaptureRequirementIfIsInstanceOfType(
                    header.cTransactionsInLog,
                    typeof(uint),
                    246,
                    @"[In Header] cTransactionsInLog (4 bytes): An unsigned integer that specifies the number of transactions in the transaction log (section 2.3.3). ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R247
            site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    header.cTransactionsInLog,
                    247,
                    @"[In Header] [cTransactionsInLog] MUST NOT be zero.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R248
            site.CaptureRequirementIfIsTrue(
                    header.cbLegacyExpectedFileLength.GetType() == typeof(uint) && header.cbLegacyExpectedFileLength == 0,
                    248,
                    @"[In Header] cbLegacyExpectedFileLength (4 bytes): An unsigned integer that MUST be zero, [and MUST be ignored.]");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R250
            site.CaptureRequirementIfIsTrue(
                    header.rgbPlaceholder.GetType() == typeof(ulong) && header.rgbPlaceholder == 0,
                    250,
                    @"[In Header] rgbPlaceholder (8 bytes): An unsigned integer that MUST be zero, [and MUST be ignored.]");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R252
            site.CaptureRequirementIfIsTrue(
                    header.fcrLegacyFileNodeListRoot.GetType() == typeof(FileChunkReference32) && header.fcrLegacyFileNodeListRoot.IsfcrNil(),
                    252,
                    @"[In Header] fcrLegacyFileNodeListRoot (8 bytes): A FileChunkReference32 structure that MUST be ""fcrNil"".");

            this.VerifyFileChunkReference32(header.fcrLegacyFileNodeListRoot, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R253
            site.CaptureRequirementIfIsTrue(
                    header.cbLegacyFreeSpaceInFreeChunkList.GetType() == typeof(uint) && header.cbLegacyFreeSpaceInFreeChunkList == 0,
                    253,
                    @"[In Header] cbLegacyFreeSpaceInFreeChunkList (4 bytes): An unsigned integer that MUST be zero, [and MUST be ignored.]");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R258
            site.CaptureRequirementIfIsTrue(
                    header.fHasNoEmbeddedFileObjects.GetType() == typeof(byte) && header.fHasNoEmbeddedFileObjects == 0,
                    258,
                    @"[In Header] fHasNoEmbeddedFileObjects (1 byte): An unsigned integer that MUST be zero, [and MUST be ignored.]");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R260
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidAncestor,
                    typeof(Guid),
                    260,
                    @"[In Header] guidAncestor (16 bytes): A GUID that specifies the Header.guidFile field of the table of contents file, as specified by [MS-ONE] section 2.1.15, given by the following table[Section file (.one) and Table of contents file (.onetoc2)]:");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R264
            site.CaptureRequirementIfIsInstanceOfType(
                    header.crcName,
                    typeof(uint),
                    264,
                    @"[In Header] crcName (4 bytes): An unsigned integer that specifies the CRC value (section 2.1.2) of the name of this revision store file. ");

            byte[] bytes = System.Text.Encoding.Unicode.GetBytes(fileName + "\0");
            CRC32 crc = new CRC32();
            crc.ComputeHash(bytes);
            uint crcValue = crc.CRC32Hash;

            if (extension == ".one")
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R265
                site.CaptureRequirementIfAreEqual<uint>(
                        crcValue,
                        header.crcName,
                        265,
                        @"[In Header] [crcName] The name is the Unicode representation of the file name with its extension and an additional null character at the end.");

                // If R265 is verified, R22 will be verified.
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R22
                site.CaptureRequirement(
                        22,
                        @"[In Cyclic Redundancy Check (CRC) Algorithms] The file format .one specifies:");

                // If R265 is verified, R24 will be verified.
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R24
                site.CaptureRequirement(
                      24,
                      @"[In Cyclic Redundancy Check (CRC) Algorithms] Normal representation for the polynomial is 0x04C11DB7.");

                // If R265 is verified, R25 will be verified.
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R25
                site.CaptureRequirement(
                      25,
                      @"[In Cyclic Redundancy Check (CRC) Algorithms] For the purpose of ordering, the least significant bit of the 32-bit CRC is defined to be the coefficient of the x31 term. ");

                // If R265 is verified, R26 will be verified.
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R26
                site.CaptureRequirement(
                      26,
                      @"[In Cyclic Redundancy Check (CRC) Algorithms] The 32-bit CRC register is initialized to all 1’s and once the data is processed, the CRC register is inverted. (1’s complement.)");
            }
            // If the R265 is verifed, the crc value that is calculated using the CRC algorithm, is equal crcName, so the R266 will be verified.
            //  Verify MS-ONESTORE requirement: MS-ONESTORE_R266
            site.CaptureRequirement(
                    266,
                    @"[In Header]  [crcName] This CRC is always calculated using the CRC algorithm for the .one file (section 2.1.2), regardless of this revision store file format.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R267
            site.CaptureRequirementIfIsInstanceOfType(
                    header.fcrHashedChunkList,
                    typeof(FileChunkReference64x32),
                    267,
                    @"[In Header] fcrHashedChunkList (12 bytes): A FileChunkReference64x32 structure (section 2.2.4.4) that specifies a reference to the first FileNodeListFragment in a hashed chunk list (section 2.3.4).");

            this.VerifyFileChunkReference64x32(header.fcrHashedChunkList, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R270
            site.CaptureRequirementIfIsInstanceOfType(
                    header.fcrTransactionLog,
                    typeof(FileChunkReference64x32),
                    270,
                    @"[In Header] fcrTransactionLog (12 bytes): A FileChunkReference64x32 structure that specifies a reference to the first TransactionLogFragment structure (section 2.3.3.1) in a transaction log (section 2.3.3).");

            this.VerifyFileChunkReference64x32(header.fcrTransactionLog, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R271
            site.CaptureRequirementIfIsTrue(
                    header.fcrTransactionLog.IsfcrNil() == false && header.fcrTransactionLog.IsfcrZero() == false,
                    271,
                    @"[In Header] The value of the fcrTransactionLog field MUST NOT be ""fcrZero"" and MUST NOT be ""fcrNil"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R272
            site.CaptureRequirementIfIsInstanceOfType(
                    header.fcrFileNodeListRoot,
                    typeof(FileChunkReference64x32),
                    272,
                    @"[In Header] fcrFileNodeListRoot (12 bytes): A FileChunkReference64x32 structure that specifies a reference to a root file node list (section 2.1.14).");

            this.VerifyFileChunkReference64x32(header.fcrFileNodeListRoot, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R273
            site.CaptureRequirementIfIsTrue(
                    header.fcrFileNodeListRoot.IsfcrNil() == false && header.fcrFileNodeListRoot.IsfcrZero() == false,
                    273,
                    @"[In Header] The value of the fcrFileNodeListRoot field MUST NOT be ""fcrZero"" and MUST NOT be ""fcrNil"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R274
            site.CaptureRequirementIfIsInstanceOfType(
                    header.fcrFreeChunkList,
                    typeof(FileChunkReference64x32),
                    274,
                    @"[In Header] fcrFreeChunkList (12 bytes): A FileChunkReference64x32 structure that specifies a reference to the first FreeChunkListFragment structure (section 2.3.2.1).");

            this.VerifyFileChunkReference64x32(header.fcrFreeChunkList, site);

            if (header.fcrFreeChunkList.IsfcrNil() || header.fcrFreeChunkList.IsfcrZero())
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R275
                site.CaptureRequirementIfAreEqual<int>(
                        0,
                        file.FreeChunkList.Count,
                        275,
                        @"[In Header] If the value of the FileChunkReference64x32 structure is ""fcrZero"" or ""fcrNil"", then the free chunk list (section 2.3.2) does not exist.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R277
            site.CaptureRequirementIfIsInstanceOfType(
                    header.cbExpectedFileLength,
                    typeof(ulong),
                    277,
                    @"[In Header] cbExpectedFileLength (8 bytes): An unsigned integer that specifies the size, in bytes, of this revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R279
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidFileVersion,
                    typeof(Guid),
                    279,
                    @"[In Header] guidFileVersion (16 bytes): A GUID, as specified by [MS-DTYP].");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R282
            site.CaptureRequirementIfIsInstanceOfType(
                    header.nFileVersionGeneration,
                    typeof(ulong),
                    282,
                    @"[In Header] nFileVersionGeneration (8 bytes): An unsigned integer that specifies the number of times the file has changed.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R284
            site.CaptureRequirementIfIsInstanceOfType(
                    header.guidDenyReadFileVersion,
                    typeof(Guid),
                    284,
                    @"[In Header] guidDenyReadFileVersion (16 bytes): A GUID, as specified by [MS-DTYP]. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R286
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    header.grfDebugLogFlags,
                    286,
                    @"[In Header] grfDebugLogFlags (4 bytes): MUST be zero.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R288
            site.CaptureRequirementIfIsTrue(
                    header.fcrDebugLog.GetType() == typeof(FileChunkReference64x32) && header.fcrDebugLog.IsfcrZero(),
                    288,
                    @"[In Header] fcrDebugLog (12 bytes): A FileChunkReference64x32 structure that MUST have a value ""fcrZero"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R188
            site.CaptureRequirementIfIsTrue(
                    header.fcrDebugLog.Cb==uint.MinValue && header.fcrDebugLog.Stp==ulong.MinValue,
                    188,
                    @"[In File Chunk Reference] fcrZero: Specifies a file chunk reference where all bits of the stp and cb fields are set to zero.");

            this.VerifyFileChunkReference64x32(header.fcrDebugLog, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R290
            site.CaptureRequirementIfIsTrue(
                    header.fcrAllocVerificationFreeChunkList.GetType() == typeof(FileChunkReference64x32) && header.fcrAllocVerificationFreeChunkList.IsfcrZero(),
                    290,
                    @"[In Header] fcrAllocVerificationFreeChunkList (12 bytes): A FileChunkReference64x32 structure that MUST be ""fcrZero"".");

            this.VerifyFileChunkReference64x32(header.fcrAllocVerificationFreeChunkList, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R292
            site.CaptureRequirementIfIsInstanceOfType(
                    header.bnCreated,
                    typeof(uint),
                    292,
                    @"[In Header] bnCreated (4 bytes): An unsigned integer that specifies the build number of the application that created this revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R294
            site.CaptureRequirementIfIsInstanceOfType(
                    header.bnLastWroteToThisFile,
                    typeof(uint),
                    294,
                    @"[In Header] bnLastWroteToThisFile (4 bytes): An unsigned integer that specifies the build number of the application that last wrote to this revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R296
            site.CaptureRequirementIfIsInstanceOfType(
                    header.bnOldestWritten,
                    typeof(uint),
                    296,
                    @"[In Header] bnOldestWritten (4 bytes): An unsigned integer that specifies the build number of the oldest application that wrote to this revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R298
            site.CaptureRequirementIfIsInstanceOfType(
                    header.bnNewestWritten,
                    typeof(uint),
                    298,
                    @"[In Header] bnNewestWritten (4 bytes): An unsigned integer that specifies the build number of the newest application that wrote to this revision store file.");

            bool isZeroOfReserved = false;
            foreach(byte b in header.rgbReserved)
            {
                if (b != 0)
                {
                    isZeroOfReserved = false;
                    break;
                }
                else
                {
                    isZeroOfReserved = true;
                }
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R300
            site.CaptureRequirementIfIsTrue(
                    isZeroOfReserved,
                    300,
                    @"[In Header] rgbReserved (728 bytes): MUST be zero.");

            if (extension == ".one")
            {
                this.VerifyHeaderInOne(header, site);
            }
            if (extension == ".onetoc2")
            {
                this.VerifyHeaderInOnetoc2(header, site);
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with Header structure in .one file.
        /// </summary>
        /// <param name="header">The instacne of header structure.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyHeaderInOne(Header header, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R221
            site.CaptureRequirementIfAreEqual<Guid>(
                    Guid.Parse("{7B5C52E4-D88C-4DA7-AEB1-5378D02996D3}"),
                    header.guidFileType,
                    221,
                    @"[In Header] File format .one's value: {7B5C52E4-D88C-4DA7-AEB1-5378D02996D3}");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R230
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000002A,
                    header.ffvLastCodeThatWroteToThisFile,
                    230,
                    @"[In Header] [ffvLastCodeThatWroteToThisFile] file format .one: 0x0000002A");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R234
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000002A,
                    header.ffvOldestCodeThatHasWrittenToThisFile,
                    234,
                    @"[In Header] [ffvOldestCodeThatHasWrittenToThisFile] File format .one: 0x0000002A");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R238
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000002A,
                    header.ffvNewestCodeThatHasWrittenToThisFile,
                    238,
                    @"[In Header] [ffvNewestCodeThatHasWrittenToThisFile] File format .one: 0x0000002A");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R242
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000002A,
                    header.ffvOldestCodeThatMayReadThisFile,
                    242,
                    @"[In Header]  [ffvOldestCodeThatMayReadThisFile] File format .one: 0x0000002A");
        }
        /// <summary>
        /// This method is used to verify the requirements related with Header structure in .onetoc file.
        /// </summary>
        /// <param name="header">The instacne of header structure.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyHeaderInOnetoc2(Header header, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R222
            site.CaptureRequirementIfAreEqual<Guid>(
                    Guid.Parse("{43FF2FA1-EFD9-4C76-9EE2-10EA5722765F}"),
                    header.guidFileType,
                    222,
                    @"[In Header] File format .onetoc2's value:{43FF2FA1-EFD9-4C76-9EE2-10EA5722765F}");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R231
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000001B,
                    header.ffvLastCodeThatWroteToThisFile,
                    231,
                    @"[In Header] [ffvLastCodeThatWroteToThisFile] file format .onetoc2: 0x0000001B");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R235
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000001B,
                    header.ffvOldestCodeThatHasWrittenToThisFile,
                    235,
                    @"[In Header] [ffvOldestCodeThatHasWrittenToThisFile] File Format .onetoc2: 0x0000001B");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R239
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000001B,
                    header.ffvNewestCodeThatHasWrittenToThisFile,
                    239,
                    @"[In Header] [ffvNewestCodeThatHasWrittenToThisFile] File Format .onetoc2: 0x0000001B");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R243
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0000001B,
                    header.ffvOldestCodeThatMayReadThisFile,
                    243,
                    @"[In Header]  [ffvOldestCodeThatMayReadThisFile] File Format .onetoc2: 0x0000001B");
        }
        private void VerifyFreeChunkList(List<FreeChunkListFragment> freeChunkList, ITestSite site)
        {

        }
        /// <summary>
        /// This method is used to verify the requirements related with Transaction Log structure.
        /// </summary>
        /// <param name="file">The instance of revision-based format file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyTransactionLog(OneNoteRevisionStoreFile file, ITestSite site)
        {
            List<TransactionLogFragment> transactionLog = file.TransactionLog;

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R321
            site.CaptureRequirement(
                321,
                @"[In Transaction Log] The TransactionEntry structures for all transactions are stored sequentially in TransactionLogFragment.sizeTable.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R327
            site.CaptureRequirement(
                327,
                @"[In  Transaction Log] The Header.fcrTransactionLog field (section 2.3.1) references the first TransactionLogFragment structure (section 2.3.3.1) in the log.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R328
            site.CaptureRequirement(
                328,
                @"[In TransactionLogFragment] The TransactionLogFragment structure specifies an array of TransactionEntry structures (section 2.3.3.2) and a reference to the next TransactionLogFragment structure if it exists.");

            List<TransactionEntry> transactionEntrys = new List<TransactionEntry>();
            uint count = 0;
            foreach(TransactionLogFragment logFragment in transactionLog)
            {
                List<byte> transactionBuffer = new List<byte>();
                foreach(TransactionEntry entry in logFragment.sizeTable)
                {
                    if(entry.srcID!=0)
                    {
                        if(entry.srcID== 0x00000001)
                        {
                            count++;
                             // Verify MS-ONESTORE requirement: MS-ONESTORE_R339
                            site.CaptureRequirementIfAreEqual<uint>(
                                    0x00000001,
                                    entry.srcID,
                                    339,
                                    @"[In TransactionEntry] A value of 0x00000001 specifies the sentinel entry. ");

                        }
                        else
                        {
  
                        }

                        transactionEntrys.Add(entry);
                    }

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R337
                    site.CaptureRequirementIfIsInstanceOfType(
                            entry.srcID,
                            typeof(uint),
                            337,
                            @"[In TransactionEntry] srcID (4 bytes): An unsigned integer that specifies the identity of the file node list modified by this transaction, or the sentinel entry for the transaction.");

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R341
                    site.CaptureRequirementIfIsInstanceOfType(
                            entry.TransactionEntrySwitch,
                            typeof(uint),
                            341,
                            @"[In TransactionEntry] TransactionEntrySwitch (4 bytes): An unsigned integer of 4 bytes in size.");
                }

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R330
                site.CaptureRequirementIfIsInstanceOfType(
                        logFragment.sizeTable,
                        typeof(TransactionEntry[]),
                        330,
                        @"[In TransactionLogFragment] sizeTable (variable): An array of TransactionEntry structures.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R332
                site.CaptureRequirementIfIsInstanceOfType(
                        logFragment.nextFragment,
                        typeof(FileChunkReference64x32),
                        332,
                        @"[In TransactionLogFragment] nextFragment (12 bytes): A FileChunkReference64x32 structure (section 2.2.4.4) that specifies the location and size of the next TransactionLogFragment structure.");

                this.VerifyFileChunkReference64x32(logFragment.nextFragment, site);
            }

            TransactionEntry lastEntry = transactionEntrys[transactionEntrys.Count - 1];

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R322
            site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                lastEntry.srcID,
                322,
                @"[In  Transaction Log] The last entry for a transaction MUST be a special sentinel entry with the value of the TransactionEntry.srcID field set to 0x00000001.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R323
            site.CaptureRequirementIfAreEqual<uint>(
                file.Header.cTransactionsInLog,
                count,
                323,
                @"[In Transaction Log] A Header.cTransactionsInLog field (section 2.3.1) that maintains the total number of transactions that have occurred. ");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R331
            site.CaptureRequirement(
                331,
                @"[In TransactionLogFragment] A transaction MUST add all of its entries to the array sequentially and MUST terminate with a sentinel entry with TransactionEntry.srcID set to 0x00000001.");

        }
        /// <summary>
        /// This method is used to verify the requirements related with Hasked Chunk List structure.
        /// </summary>
        /// <param name="hashedChunkList">The insatance of Hasked Chunk List structure.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyHashedChunkList(List<FileNodeListFragment> hashedChunkList, ITestSite site)
        {
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R351
            site.CaptureRequirement(
                     351,
                     @"[In Hashed Chunk List] The Header.fcrHashedChunkList field (section 2.3.1) references the first FileNodeListFragment structure (section 2.4.1) in the hashed chunk list, if it exists.");

            foreach(FileNodeListFragment fragment in hashedChunkList)
            {
                foreach(FileNode node in fragment.rgFileNodes)
                {
                    if (node.FileNodeID != 0)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R350
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x0C2,
                                (uint)node.FileNodeID,
                                350,
                                @"[In Hashed Chunk List] A hashed chunk list is an optional file node list (section 2.4) that specifies a collection of FileNode structures (section 2.4.3) with FileNodeID field values equal to ""0x0C2"" (HashedChunkDescriptor2FND structure, section 2.3.4.1).");

                        HashedChunkDescriptor2FND fnd = node.fnd as HashedChunkDescriptor2FND;

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R354
                        site.CaptureRequirementIfIsInstanceOfType(
                                fnd.BlobRef,
                                typeof(FileNodeChunkReference),
                                354,
                                @"[In HashedChunkDescriptor2FND] BlobRef (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies the location and size of an ObjectSpaceObjectPropSet structure.");

                        this.VerifyFileNodeChunkReference(fnd.BlobRef, site);

                        // If the OneNote file parse successfully, this requirement will be verified.
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R355
                        site.CaptureRequirement(
                                355,
                                @"[In  HashedChunkDescriptor2FND] guidHash (16 bytes): An unsigned integer that specifies an MD5 checksum, as specified in [RFC1321], of data referenced by the BlobRef field.");
                        }
                }
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with Root File Node List.
        /// </summary>
        /// <param name="rootFileNodeList">The insatance of Root File Node List.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRootFileNodeList(RootFileNodeList rootFileNodeList, string fileName, ITestSite site)
        {
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R159
            site.CaptureRequirement(
                    159,
                    @"[In Root File Node List] The root file node list MUST begin with the FileNodeListFragment structure (section 2.4.1) specified by the Header.fcrFileNodeListRoot field (section 2.3.1).");

            bool isFileNodeVerified = false;
            foreach (FileNode node in rootFileNodeList.FileNodeSequence)
            {
                isFileNodeVerified = true;
                if ((uint)node.FileNodeID != 0x008 && (uint)node.FileNodeID != 0x004 && (uint)node.FileNodeID != 0x090)
                {
                    isFileNodeVerified = false;
                    break;
                }
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R160
            site.CaptureRequirementIfIsTrue(
                    isFileNodeVerified,
                    160,
                    @"[In Root File Node List] The root file node list MUST consist of the following FileNode structures[FileNode structures with FileNodeID field values equal to 0x008,FileNode structure with a FileNodeID field value equal to 0x004,FileNode structure with FileNodeID field values equal to 0x090  ] (section 2.4.3), and MUST NOT contain any others:");

            FileNode[] objectSpaceManifestListReferences = rootFileNodeList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectSpaceManifestListReferenceFND).ToArray();

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R161
            site.CaptureRequirementIfIsTrue(
                    objectSpaceManifestListReferences.Length >= 1,
                    161,
                    @"[In Root File Node List] • One or more FileNode structures with FileNodeID field values equal to 0x008 (ObjectSpaceManifestListReferenceFND structure, section 2.5.2).");

            // If R161 is verified, then R43 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R43
            site.CaptureRequirement(
                    43,
                    @"[In Object Space] Object spaces MUST be referenced from the root file node list (section 2.1.14) by a FileNode structure (section 2.4.3) with a FileNodeID field value equal to 0x08 (ObjectSpaceManifestListReferenceFND structure, section 2.5.2).");

            for (int i=0;i<objectSpaceManifestListReferences.Length;i++)
            {
                ObjectSpaceManifestListReferenceFND fnd = objectSpaceManifestListReferences[i].fnd as ObjectSpaceManifestListReferenceFND;
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R502
                site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                        fnd.gosid,
                        ((ObjectSpaceManifestListStartFND)rootFileNodeList.ObjectSpaceManifestList[i].FileNodeSequence[0].fnd).gosid,
                        502,
                        @"[In ObjectSpaceManifestListStartFND] [gosid] MUST match the ObjectSpaceManifestListReferenceFND.gosid field (section 2.5.2) of the FileNode structure that referenced this file node list (section 2.4).");

                // If R502 is verified,then R44 will be verified.
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R44
                site.CaptureRequirement(
                        44,
                        @"[In Object Space] Object spaces MUST have a unique identifier (OSID), specified by the ObjectSpaceManifestListReferenceFND.gosid field.");

                this.VerifyExtendedGUID(fnd.gosid, site);
                this.VerifyObjectSpaceManifestListReferenceFND(fnd, site);
            }

            FileNode[] objectSpaceManifestRoots = rootFileNodeList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectSpaceManifestRootFND).ToArray();

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R163
            site.CaptureRequirementIfIsTrue(
                    objectSpaceManifestRoots.Length == 1,
                    163,
                    @"[In Root File Node List] • One FileNode structure with a FileNodeID field value equal to 0x004 (ObjectSpaceManifestRootFND structure, section 2.5.1).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R489
            site.CaptureRequirementIfAreEqual<int>(
                    1,
                    objectSpaceManifestRoots.Length,
                    489,
                    @"[In ObjectSpaceManifestRootFND] There MUST be only one ObjectSpaceManifestRootFND structure (section 2.5.1) in the revision store file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R490
            site.CaptureRequirementIfIsNotNull(
                    objectSpaceManifestRoots[0],
                    490,
                    @"[In ObjectSpaceManifestRootFND] This FileNode structure MUST be in the root file node list (section 2.1.14).");


            this.VerifyObjectSpaceManifestRootFND((ObjectSpaceManifestRootFND)objectSpaceManifestRoots[0].fnd, (ObjectSpaceManifestListReferenceFND)objectSpaceManifestListReferences[0].fnd, site);

            FileNode[] fileDataStoreListReferences = rootFileNodeList.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.FileDataStoreListReferenceFND).ToArray();

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R164
            site.CaptureRequirementIfIsTrue(
                    fileDataStoreListReferences.Length >= 0,
                    164,
                    @"[In Root File Node List] • Zero or one FileNode structure with FileNodeID field values equal to 0x090 (FileDataStoreListReferenceFND structure, section 2.5.21).");

            foreach (FileNode node in fileDataStoreListReferences)
            {
                this.VerifyFileDataStoreListReferenceFND((FileDataStoreListReferenceFND)node.fnd, site);
            }

            foreach (ObjectSpaceManifestList objSpaceManifestList in rootFileNodeList.ObjectSpaceManifestList)
            {
                this.VerifyObjectSpaceManifestList(objSpaceManifestList, fileName, site);
            }
            this.VerifyFileNodeList(rootFileNodeList.FileNodeListFragments, fileName, site);
        }

        /// <summary>
        /// This method is used to verify the requirements related with Object Space Manifest List.
        /// </summary>
        /// <param name="instance">The insatance of Object Space Manifest List.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectSpaceManifestList(ObjectSpaceManifestList instance, string fileName, ITestSite site)
        {
            FileNode firstNode = instance.FileNodeSequence[0];

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R81 
            site.CaptureRequirementIfAreEqual<uint>(
                    0x00C,
                    (uint)firstNode.FileNodeID,
                    81,
                    @"[In Object Space Manifest List] 1. A FileNode structure with FileNodeID field value equal to 0x00C (ObjectSpaceManifestListStartFND structure, section 2.5.3).");

            this.VerifyObjectSpaceManifestListStartFND((ObjectSpaceManifestListStartFND)firstNode.fnd, site);

            FileNode[] RevisionManifestListRefArray = instance.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.RevisionManifestListReferenceFND).ToArray();
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R82
            site.CaptureRequirementIfIsTrue(
                    RevisionManifestListRefArray.Length >= 1,
                    82,
                    @" [In Object Space Manifest List]2. One or more FileNode structures with FileNodeID field values equal to 0x010 (RevisionManifestListReferenceFND structure, section 2.5.4).");

            // If R82 is verified,then R45 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R45
            site.CaptureRequirement(
                    45,
                    @"[In Object Space] Every revision store file MUST have exactly one root object space whose OSID is specified by the ObjectSpaceManifestRootFND.gosidRoot field.");

            foreach (FileNode node in RevisionManifestListRefArray)
            {
                this.VerifyRevisionManifestListReferenceFND((RevisionManifestListReferenceFND)node.fnd, site);
            }

            foreach(RevisionManifestList revisionManifestList in instance.RevisionManifestList)
            {
                this.VerifyRevisionManifestList(revisionManifestList, fileName, site);

                ExtendedGUID gosid = ((ObjectSpaceManifestListStartFND)instance.ObjectSpaceManifestListStart.fnd).gosid;
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R122
                site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                        gosid,
                        ((RevisionManifestListStartFND)revisionManifestList.FileNodeSequence[0].fnd).gosid,
                        122,
                        @"[In Revision Manifest List] All of the revision manifests (section 2.1.9) for an object space MUST appear in a single revision manifest list. ");

                this.VerifyExtendedGUID(gosid, site);
            }
            this.VerifyFileNodeList(instance.FileNodeListFragments, fileName, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with Revision Manifest List.
        /// </summary>
        /// <param name="instance">The insatance of Revision Manifest List.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestList(RevisionManifestList instance, string fileName, ITestSite site)
        {
            FileNode firstNode = instance.FileNodeSequence[0];

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R125
            site.CaptureRequirementIfAreEqual<uint>(
                    0x014,
                    (uint)firstNode.FileNodeID,
                    125,
                    @"[In Revision Manifest List] The first FileNode structure (section 2.4.3) in a revision manifest list MUST have the FileNodeID field value equal to 0x14 (RevisionManifestListStartFND structure, section 2.5.5).");

            this.VerifyRevisionManifestListStartFND((RevisionManifestListStartFND)firstNode.fnd, site);

            FileNode[] fileNodes = instance.FileNodeSequence.Where(f => f.FileNodeID != FileNodeIDValues.RevisionManifestListStartFND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionRoleDeclarationFND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionRoleAndContextDeclarationFND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestStart6FND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestStart7FND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestStart4FND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectGroupListReferenceFND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectInfoDependencyOverridesFND &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableStart2FND &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEntryFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEndFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.RootObjectReference3FND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectInfoDependencyOverridesFND &&
                                                                  f.FileNodeID != FileNodeIDValues.RootObjectReference2FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableStartFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEntry2FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEntry3FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.DataSignatureGroupDefinitionFND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectDeclarationWithRefCountFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectRevisionWithRefCountFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectRevisionWithRefCount2FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestEndFND &&
                                                                  f.FileNodeID!=FileNodeIDValues.ObjectDataEncryptionKeyV2FNDX).ToArray();

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R126
            site.CaptureRequirementIfIsTrue(
                    fileNodes.Length==0,
                    126,
                    @"[In Revision Manifest List] The remainder of the revision manifest list MUST contain zero or more of the following structures, and MUST NOT contain any others:
• Revision manifests (section 2.1.9).
• FileNode structures with a FileNodeID field value equal to 0x5C (RevisionRoleDeclarationFND structure, section 2.5.17).
• FileNode structures with a FileNodeID field value equal to 0x5D (RevisionRoleAndContextDeclarationFND structure, section 2.5.18).");

            foreach(FileNode revisionRoleDeclarationNode in instance.RevisionRoleDeclaration)
            {
                this.VerifyRevisionRoleDeclarationFND((RevisionRoleDeclarationFND)revisionRoleDeclarationNode.fnd, site);
            }

            foreach(FileNode RevisionRoleAndContextDeclaration in instance.RevisionRoleAndContextDeclaration)
            {
                this.VerifyRevisionRoleAndContextDeclarationFND((RevisionRoleAndContextDeclarationFND)RevisionRoleAndContextDeclaration.fnd, site);
            }

            foreach(RevisionManifest revisionmanifest in instance.RevisionManifests)
            {
                this.VerifyRevisionManifest(revisionmanifest, fileName, site);
            }

            foreach(ObjectGroupList objectGroup in instance.ObjectGroupList)
            {
                this.VerifyObjectGroupList(objectGroup,fileName, site);
            }

            FileNode[] objectGroupRefArray = instance.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.ObjectGroupListReferenceFND).ToArray();
            for (int i = 0; i < objectGroupRefArray.Length; i++)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R720
                site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                        ((ObjectGroupStartFND)instance.ObjectGroupList[i].FileNodeSequence[0].fnd).oid,
                        ((ObjectGroupListReferenceFND)objectGroupRefArray[i].fnd).ObjectGroupID,
                        720,
                        @"[In ObjectGroupListReferenceFND] [ObjectGroupID] MUST be the same value as the ObjectGroupStartFND.oid field value of the object group that the ref field points to.");

            }
            this.VerifyFileNodeList(instance.FileNodeListFragments, fileName, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with Revision Manifest List.
        /// </summary>
        /// <param name="instance">The instance of Revision Manifest.</param>
        /// <param name="fileName">The file name of the OneNote file.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifest(RevisionManifest instance, string fileName, ITestSite site)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            FileNode firstNode = instance.FileNodeSequence[0];

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R100
            site.CaptureRequirementIfIsTrue(
                    (extension == ".one" && ((uint)firstNode.FileNodeID == 0x01E || (uint)firstNode.FileNodeID == 0x01F) ||
                    (extension == ".onetoc2" && (uint)firstNode.FileNodeID == 0x01B)),
                    100,
                    @"[In Revision Manifest] The sequence MUST begin with one of the FileNode structures[0x01E,0x01F in .one and 0x01B in .onetoc2] described in the following table:");

            if (extension == ".one")
            {
                if (firstNode.FileNodeID == FileNodeIDValues.RevisionManifestStart6FND)
                {
                    this.VerifyRevisionManifestStart6FND((RevisionManifestStart6FND)firstNode.fnd, site);
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R101
                    site.CaptureRequirementIfAreEqual<uint>(
                            0x01E,
                            (uint)firstNode.FileNodeID,
                            101,
                            @"[In Revision Manifest] file format .one: 0x01E (RevisionManifestStart6FND structure, section 2.5.7)");
                }
                if (firstNode.FileNodeID == FileNodeIDValues.RevisionManifestStart7FND)
                {
                    this.VerifyRevisionManifestStart7FND((RevisionManifestStart7FND)firstNode.fnd, site);
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R102
                    site.CaptureRequirementIfAreEqual<uint>(
                            0x01F,
                            (uint)firstNode.FileNodeID,
                            102,
                            @"[In Revision Manifest] file format .one: 0x01F (RevisionManifestStart7FND structure, section 2.5.8)");
                }

                FileNode[] nodes = instance.FileNodeSequence.Where(f => f.FileNodeID != FileNodeIDValues.RevisionManifestStart6FND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestStart7FND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectDataEncryptionKeyV2FNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectGroupListReferenceFND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectInfoDependencyOverridesFND &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableStart2FND &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEntryFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.GlobalIdTableEndFNDX &&
                                                                  f.FileNodeID != FileNodeIDValues.RootObjectReference3FND &&
                                                                  f.FileNodeID != FileNodeIDValues.ObjectInfoDependencyOverridesFND &&
                                                                  f.FileNodeID != FileNodeIDValues.RevisionManifestEndFND).ToArray();

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R105
                site.CaptureRequirementIfIsTrue(
                        nodes.Length == 0,
                        105,
                        @"[In Revision Manifest] The remainder of the sequence can contain the FileNode structures described in the following table, and MUST NOT contain any other FileNode structures.");

                for (int i = 0; i < instance.FileNodeSequence.Count; i++)
                {
                    FileNode node = instance.FileNodeSequence[i];
                    if (node.FileNodeID == FileNodeIDValues.ObjectGroupListReferenceFND)
                    {
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R106
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x0B0,
                                (uint)node.FileNodeID,
                                106,
                                @"[In Revision Manifest] [File format] .one: Zero or more sequences of object group FileNode structures:
§ 0x0B0 (ObjectGroupListReferenceFND structure, section 2.5.31)");

                        // If R106 is verifed, so R717 will be verified.
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R717
                        site.CaptureRequirement(
                                717,
                                @"[In ObjectGroupListReferenceFND] The value of the FileNode.FileNodeID field MUST be set to 0x0B0.");

                        this.VerifyObjectGroupListReferenceFND((ObjectGroupListReferenceFND)node.fnd, site);

                        FileNode nextNode = instance.FileNodeSequence[i + 1];

                        site.CaptureRequirementIfAreEqual<FileNodeIDValues>(
                                FileNodeIDValues.ObjectInfoDependencyOverridesFND,
                                nextNode.FileNodeID,
                                108,
                                @"[In Revision Manifest] [File format] .one: Where each ObjectGroupListReferenceFND structure MUST be followed by an ObjectInfoDependencyOverridesFND structure.");

                        this.VerifyObjectInfoDependencyOverridesFND((ObjectInfoDependencyOverridesFND)nextNode.fnd, site);
                        i = i + 1;
                    }
                    if (node.FileNodeID == FileNodeIDValues.GlobalIdTableStart2FND)
                    {
                        List<FileNode> globalIdTable = new List<FileNode>();
                        for (int j = i; j < instance.FileNodeSequence.Count; j++)
                        {
                            globalIdTable.Add(instance.FileNodeSequence[j]);
                            if (instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEndFNDX)
                            {
                                this.VerifyGlobalIdentificationTableInOne(globalIdTable, site);
                                i = j + 1;
                                break;
                            }
                            if(instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEntryFNDX)
                            {
                                this.VerifyGlobalIdTableEntryFNDX((GlobalIdTableEntryFNDX)instance.FileNodeSequence[j].fnd, site);
                            }
                        }

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R109
                        site.CaptureRequirementIfIsTrue(
                            globalIdTable[0].FileNodeID == FileNodeIDValues.GlobalIdTableStart2FND &&
                            globalIdTable.Where(f => f.FileNodeID == FileNodeIDValues.GlobalIdTableEntryFNDX).ToArray().Length >= 0 &&
                            globalIdTable[globalIdTable.Count - 1].FileNodeID == FileNodeIDValues.GlobalIdTableEndFNDX,
                            109,
                            @"[In Revision Manifest] [File format] .one: Zero or one sequence of global identification table FileNode structures:
§ A FileNode structure with a FileNodeID field value equal to 0x022 (GlobalIdTableStart2FND structure, section 2.4.3)
§ Zero or more FileNode structures with a FileNodeID field value equal to 0x024 (GlobalIdTableEntryFNDX structure, section 2.5.10)
§ A FileNode structure with a FileNodeID field value equal to 0x028 (GlobalIdTableEndFNDX structure, section 2.4.3)");

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R112
                        site.CaptureRequirementIfIsTrue(
                                globalIdTable.Where(f => f.FileNodeID == FileNodeIDValues.ObjectGroupListReferenceFND || f.FileNodeID == FileNodeIDValues.ObjectInfoDependencyOverridesFND).ToArray().Length == 0,
                                112,
                                @"[In Revision Manifest] [File format] .one: Where the global identification table sequence of FileNode structures MUST NOT be followed by any object group sequences of FileNode structures.");
                    }

                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R113
                    site.CaptureRequirementIfIsTrue(
                            instance.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.RootObjectReference3FND || f.FileNodeID == FileNodeIDValues.ObjectInfoDependencyOverridesFND).ToArray().Length >= 0,
                            113,
                            @"[In Revision Manifest] [File format] .one: Zero or more FileNode structures with FileNodeID field values equal to any of: 
§ 0x05A (RootObjectReference3FND structure section 2.5.16)
§ 0x084 (ObjectInfoDependencyOverridesFND structure, section 2.5.20)");
                }
            }
            else if (extension == ".onetoc2")
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R103
                site.CaptureRequirementIfAreEqual<uint>(
                        0x01B,
                        (uint)firstNode.FileNodeID,
                        103,
                        @"[In Revision Manifest] file format .onetoc2: 0x01B (RevisionManifestStart4FND structure, section 2.5.6");

                this.VerifyRevisionManifestStart4FND((RevisionManifestStart4FND)firstNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R114
                site.CaptureRequirementIfIsTrue(
                        instance.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.RootObjectReference2FNDX || f.FileNodeID == FileNodeIDValues.ObjectInfoDependencyOverridesFND).ToArray().Length >= 0,
                        114,
                        @"[In Revision Manifest] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to any of the following:
§ 0x059 (RootObjectReference2FNDX structure, section 2.5.15)
§ 0x084 (ObjectInfoDependencyOverridesFND structure)");
                
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R118
                site.CaptureRequirementIfIsTrue(
                        instance.FileNodeSequence.Where(f => f.FileNodeID == FileNodeIDValues.DataSignatureGroupDefinitionFND || 
                        f.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCountFNDX ||
                        f.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX ||
                        f.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCountFNDX ||
                        f.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCount2FNDX).ToArray().Length >= 0,
                        118,
                        @"[In Revision Manifest] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to the following:
§ 0x08C (DataSignatureGroupDefinitionFND structure, section 2.5.33)
§ 0x02D (ObjectDeclarationWithRefCountFNDX structure, section 2.5.23)
§ 0x02E (ObjectDeclarationWithRefCount2FNDX structure, section 2.5.24)
§ 0x041 (ObjectRevisionWithRefCountFNDX structure, section 2.5.13)
§ 0x042 (ObjectRevisionWithRefCount2FNDX structure, section 2.5.14)");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R120
                site.CaptureRequirementIfAreEqual<uint>(
                        0x01C,
                        (uint)instance.FileNodeSequence[instance.FileNodeSequence.Count-1].FileNodeID,
                        120,
                        @"[In Revision Manifest] [File format] .onetoc2: The sequence MUST end with a FileNode structure with a FileNodeID field value equal to 0x01C (RevisionManifestEndFND structure, section 2.4.3).");
                
                int globalIdTableInde = 0;
                for (int i = 0; i < instance.FileNodeSequence.Count; i++)
                {
                    FileNode node = instance.FileNodeSequence[i];
                    if (node.FileNodeID == FileNodeIDValues.GlobalIdTableStartFNDX)
                    {
                        List<FileNode> globalIdTable = new List<FileNode>();
                        globalIdTableInde = i;
                        for (int j = i; j < instance.FileNodeSequence.Count; j++)
                        {
                            globalIdTable.Add(instance.FileNodeSequence[j]);
                            if (instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEndFNDX)
                            {
                                this.VerifyGlobalIdentificationTableInOnetoc2(globalIdTable, site);
                                i = j + 1;
                                break;
                            }
                            if(instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEntryFNDX)
                            {
                                this.VerifyGlobalIdTableEntryFNDX((GlobalIdTableEntryFNDX)instance.FileNodeSequence[j].fnd, site);
                            }
                            if (instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEntry2FNDX)
                            {
                                this.VerifyGlobalIdTableEntry2FNDX((GlobalIdTableEntry2FNDX)instance.FileNodeSequence[j].fnd, site);
                            }
                            if (instance.FileNodeSequence[j].FileNodeID == FileNodeIDValues.GlobalIdTableEntry3FNDX)
                            {
                                this.VerifyGlobalIdTableEntry3FNDX((GlobalIdTableEntry3FNDX)instance.FileNodeSequence[j].fnd, site);
                            }
                        }

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R115
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x021,
                                (uint)globalIdTable[0].FileNodeID,
                                115,
                                @"[In Revision Manifest] [File format] .onetoc2: Zero or one sequence of global identification table FileNode structures:
§ A FileNode structure with a FileNodeID field value equal to 0x021 (GlobalIdTableStartFNDX structure, section 2.5.9)");

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R116
                        site.CaptureRequirementIfIsTrue(
                                globalIdTable.Where(f=>f.FileNodeID== FileNodeIDValues.GlobalIdTableEntryFNDX ||
                                f.FileNodeID == FileNodeIDValues.GlobalIdTableEntry2FNDX ||
                                f.FileNodeID == FileNodeIDValues.GlobalIdTableEntry3FNDX).ToArray().Length>=0,
                                116,
                                @"[In Revision Manifest] [File format] .onetoc2: [Zero or one sequence of global identification table FileNode structures:] § Zero or more FileNode structures with FileNodeID field values equal to one of:
§ 0x024 (GlobalIdTableEntryFNDX structure, section 2.5.10)
§ 0x025 (GlobalIdTableEntry2FNDX structure, section 2.5.11)
§ 0x026 (GlobalIdTableEntry3FNDX structure, section 2.5.12)");

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R117
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x028,
                                (uint)globalIdTable[globalIdTable.Count-1].FileNodeID,
                                117,
                                @"[In Revision Manifest] [File format] .onetoc2: [Zero or one sequence of global identification table FileNode structures:]§ A FileNode structure with a FileNodeID field value equal to 0x028 (GlobalIdTableEndFNDX structure, section 2.4.3)");
                    }
                    if(node.FileNodeID == FileNodeIDValues.DataSignatureGroupDefinitionFND ||
                        node.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCountFNDX ||
                        node.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX ||
                        node.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCountFNDX ||
                        node.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCount2FNDX)
                    {
                        if(node.FileNodeID == FileNodeIDValues.DataSignatureGroupDefinitionFND)
                        {
                            this.VerifyDataSignatureGroupDefinitionFND((DataSignatureGroupDefinitionFND)node.fnd, site);
                        }
                        if (node.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCountFNDX)
                        {
                            this.VerifyObjectDeclarationWithRefCountFNDX((ObjectDeclarationWithRefCountFNDX)node.fnd, site);
                        }
                        if (node.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX)
                        {
                            this.VerifyObjectDeclarationWithRefCount2FNDX((ObjectDeclarationWithRefCount2FNDX)node.fnd, site);
                        }
                        if (node.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCountFNDX)
                        {
                            this.VerifyObjectRevisionWithRefCountFNDX((ObjectRevisionWithRefCountFNDX)node.fnd, site);
                        }
                        if (node.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCount2FNDX)
                        {
                            this.VerifyObjectRevisionWithRefCount2FNDX((ObjectRevisionWithRefCount2FNDX)node.fnd, site);
                        }

                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R119
                        site.CaptureRequirementIfIsTrue(
                                i> globalIdTableInde,
                                119,
                                @"[In Revision Manifest] [File format] .onetoc2: that MUST follow a global identification table sequence.");
                    }
                }
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with Object Group List.
        /// </summary>
        /// <param name="instance">The instance of Object Group List.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectGroupList(ObjectGroupList instance,string fileName, ITestSite site)
        {
            this.VerifyFileNodeList(instance.FileNodeListFragments, fileName, site);
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R150
            site.CaptureRequirement(
                    150,
                    @"[In Object Group] An object group MUST NOT be referenced by more than one revision manifest.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R151
            site.CaptureRequirement(
                    151,
                    @"[In Object Group] An object group MUST be contained within a single file node list (section 2.4).");

            FileNode firstNode = instance.FileNodeSequence[0];
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R153
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0B4,
                    (uint)firstNode.FileNodeID,
                    153,
                    @"[In Object Group] § FileNode structure (section 2.4.3) with a FileNodeID field value equal to 0x0B4 (ObjectGroupStartFND structure, section 2.5.32).");

            this.VerifyObjectGroupStartFND((ObjectGroupStartFND)firstNode.fnd, site);

            FileNode lastNode = instance.FileNodeSequence[instance.FileNodeSequence.Count - 1];
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R155
            site.CaptureRequirementIfAreEqual<uint>(
                    0x0B8,
                    (uint)lastNode.FileNodeID,
                    155,
                    @"[In Object Group] § FileNode structure with a FileNodeID field value equal to 0x0B8 (ObjectGroupEndFND structure, section 2.4.3).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R156
            site.CaptureRequirementIfIsTrue(
                    instance.FileNodeSequence.Where(f=>f.FileNodeID==FileNodeIDValues.DataSignatureGroupDefinitionFND ||
                                                       f.FileNodeID == FileNodeIDValues.ObjectDeclaration2RefCountFND ||
                                                       f.FileNodeID == FileNodeIDValues.ObjectDeclaration2LargeRefCountFND ||
                                                       f.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2RefCountFND ||
                                                       f.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2LargeRefCountFND ||
                                                       f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3RefCountFND ||
                                                       f.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3LargeRefCountFND).ToArray().Length>=0,
                    156,
                    @"[In Object Group] Zero or more FileNode structures with any of the following FileNodeID field values:
§ 0x08C (DataSignatureGroupDefinitionFND structure, section 2.5.33).
§ 0x0A4 (ObjectDeclaration2RefCountFND structure, section 2.5.25).
§ 0x0A5 (ObjectDeclaration2LargeRefCountFND structure, section 2.5.26).
§ 0x0C4 (ReadOnlyObjectDeclaration2RefCountFND structure, section 2.5.29).
§ 0x0C5 (ReadOnlyObjectDeclaration2LargeRefCountFND structure, section 2.5.30).
§ 0x072 (ObjectDeclarationFileData3RefCountFND structure, section 2.5.27).
§ 0x073 (ObjectDeclarationFileData3LargeRefCountFND structure, section 2.5.28).");

            foreach(FileNode node in instance.FileNodeSequence)
            {
                switch (node.FileNodeID)
                {
                    case FileNodeIDValues.DataSignatureGroupDefinitionFND:
                        this.VerifyDataSignatureGroupDefinitionFND((DataSignatureGroupDefinitionFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ObjectDeclaration2RefCountFND:
                        this.VerifyObjectDeclaration2RefCountFND((ObjectDeclaration2RefCountFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ObjectDeclaration2LargeRefCountFND:
                        this.VerifyObjectDeclaration2LargeRefCountFND((ObjectDeclaration2LargeRefCountFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ReadOnlyObjectDeclaration2RefCountFND:
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R704
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x0C4,
                                (uint)node.FileNodeID,
                                704,
                                @"[In ReadOnlyObjectDeclaration2RefCountFND] The value of the FileNode.FileNodeID field MUST be 0x0C4.");

                        this.VerifyReadOnlyObjectDeclaration2RefCountFND((ReadOnlyObjectDeclaration2RefCountFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ReadOnlyObjectDeclaration2LargeRefCountFND:
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R711
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x0C5,
                                (uint)node.FileNodeID,
                                711,
                                @"[In ReadOnlyObjectDeclaration2LargeRefCountFND] The value of the FileNode.FileNodeID field MUST be 0x0C5.");

                        this.VerifyReadOnlyObjectDeclaration2LargeRefCountFND((ReadOnlyObjectDeclaration2LargeRefCountFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ObjectDeclarationFileData3RefCountFND:
                        this.VerifyObjectDeclarationFileData3RefCountFND((ObjectDeclarationFileData3RefCountFND)node.fnd, site);
                        break;
                    case FileNodeIDValues.ObjectDeclarationFileData3LargeRefCountFND:
                        // Verify MS-ONESTORE requirement: MS-ONESTORE_R684
                        site.CaptureRequirementIfAreEqual<uint>(
                                0x073,
                                (uint)node.FileNodeID,
                                684,
                                @"[In ObjectDeclarationFileData3LargeRefCountFND] The value of the FileNode.FileNodeID field MUST be 0x073.");

                        this.VerifyObjectDeclarationFileData3LargeRefCountFND((ObjectDeclarationFileData3LargeRefCountFND)node.fnd, site);
                        break;
                }
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with global identification table in .one file.
        /// </summary>
        /// <param name="globalIdentificationTable">The global identification table FileNodes</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdentificationTableInOne(List<FileNode> globalIdentificationTable, ITestSite site)
        {
            FileNode firstNode = globalIdentificationTable[0];
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R32
            site.CaptureRequirementIfAreEqual<uint>(
                    0x022,
                    (uint)firstNode.FileNodeID,
                    32,
                    @"[In Global Identification Table] [File format] .one: A FileNode structure (section 2.4.3) with a FileNodeID field value equal to 0x022 (GlobalIdTableStart2FND structure, section 2.4.3).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R33
            site.CaptureRequirementIfIsTrue(
                    globalIdentificationTable.Where(f=>f.FileNodeID==FileNodeIDValues.GlobalIdTableEntryFNDX).ToArray().Length>=0,
                    33,
                    @"[In Global Identification Table] [File format] .one: Zero or more FileNode structures with FileNodeID field value equal to 0x024 (GlobalIdTableEntryFNDX structure, section 2.5.10).");

            FileNode lastNode = globalIdentificationTable[globalIdentificationTable.Count - 1];
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R34
            site.CaptureRequirementIfAreEqual<uint>(
                    0x028,
                    (uint)lastNode.FileNodeID,
                    34,
                    @"[In Global Identification Table] [File format] .one: A FileNode structure with a FileNodeID field value equal to 0x028 (GlobalIdTableEndFNDX structure, section 2.4.3).");
        }
        /// <summary>
        /// This method is used to verify the requirements related with global identification table in .onetoc2 file.
        /// </summary>
        /// <param name="globalIdentificationTable">The global identification table FileNodes</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdentificationTableInOnetoc2(List<FileNode> globalIdentificationTable, ITestSite site)
        {
            FileNode firstNode = globalIdentificationTable[0];
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R35
            site.CaptureRequirementIfAreEqual<uint>(
                    0x021,
                    (uint)firstNode.FileNodeID,
                    35,
                    @"[In Global Identification Table] [File format] .onetoc2:  A FileNode structure (section 2.4.3) with a FileNodeID field value equal to 0x021 (GlobalIdTableStartFNDX structure, section 2.5.9).");

            // If R35 is verified, so R541 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R541
            site.CaptureRequirement(
                    541,
                    @"[In GlobalIdTableStartFNDX] The value of the FileNode.FileNodeID field MUST be 0x021.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R36
            site.CaptureRequirementIfIsTrue(
                    globalIdentificationTable.Where(f=>f.FileNodeID==FileNodeIDValues.GlobalIdTableEntryFNDX).ToArray().Length>=0,
                    36,
                    @"[In Global Identification Table] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to 0x024 (GlobalIdTableEntryFNDX structure, section 2.5.10).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R37
            site.CaptureRequirementIfIsTrue(
                    globalIdentificationTable.Where(f => f.FileNodeID == FileNodeIDValues.GlobalIdTableEntry2FNDX).ToArray().Length >= 0,
                    37,
                    @"[In Global Identification Table] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to 0x025 (GlobalIdTableEntry2FNDX structure, section 2.5.11).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R38
            site.CaptureRequirementIfIsTrue(
                    globalIdentificationTable.Where(f => f.FileNodeID == FileNodeIDValues.GlobalIdTableEntry3FNDX).ToArray().Length >= 0,
                    38,
                    @"[In Global Identification Table] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to 0x026 (GlobalIdTableEntry3FNDX structure, section 2.5.12).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R39
            site.CaptureRequirementIfAreEqual<uint>(
                    0x028,
                    (uint)globalIdentificationTable[globalIdentificationTable.Count - 1].FileNodeID,
                    39,
                    @"[In Global Identification Table] [File format] .onetoc2: Zero or more FileNode structures with FileNodeID field values equal to 0x026 (GlobalIdTableEntry3FNDX structure, section 2.5.12).");
        }

        /// <summary>
        /// This method is used to verify the requirements related with File Node List structure.
        /// </summary>
        /// <param name="fileNodeList">The array of File Node List Fragment.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileNodeList(List<FileNodeListFragment> fileNodeList, string fileName, ITestSite site)
        {
            string extension = Path.GetExtension(fileName).ToLowerInvariant();
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R362
            site.CaptureRequirement(
                    362,
                    @"[In File Node List] All file node list fragments in a file MUST form a tree.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R363
            site.CaptureRequirement(
                    363,
                    @"[In File Node List] The Header.fcrFileNodeListRoot field (section 2.3.1) specifies the first fragment of the file node list that is the root of the tree.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R365
            site.CaptureRequirement(
                    365,
                    @"[In FileNodeListFragment] The size of the FileNodeListFragment structure is specified by the structure that references it.");

            foreach(FileNodeListFragment fileNodeFragment in fileNodeList)
            {
                this.VerifyFileNodeListFragment(fileNodeFragment, extension, site);
            }

            FileNode lastFileNode = fileNodeList[fileNodeList.Count - 1].rgFileNodes[fileNodeList[fileNodeList.Count - 1].rgFileNodes.Count - 1];
            // Verify MS-ONESTORE requirement: MS - ONESTORE_R487
            site.CaptureRequirementIfAreNotEqual<uint>(
                    0x0FF,
                    (uint)lastFileNode.FileNodeID,
                    487,
                    @"[In FileNode][FileNodeID value 0x0FF] MUST NOT be used in FileNodeListFragment structure that is the last fragment in the containing file node list.");

            FileNodeListHeader firstHeader = fileNodeList[0].Header;

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R393
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    firstHeader.nFragmentSequence,
                    393,
                    @"[In FileNodeListHeader] The nFragmentSequence field of the first fragment in a given file node list MUST be 0 and the nFragmentSequence fields of all subsequent fragments in this list MUST be sequential.");

        }
        /// <summary>
        /// This method is used to verify the requirements related with FileNodeListFragment structure.
        /// </summary>
        /// <param name="fileNodeListFragment">The instance of FileNodeListFragment.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileNodeListFragment(FileNodeListFragment fileNodeListFragment,string extension, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R368
            site.CaptureRequirementIfIsInstanceOfType(
                fileNodeListFragment.Header,
                typeof(FileNodeListHeader),
                368,
                @"[In FileNodeListFragment] header (16 bytes): A FileNodeListHeader structure (section 2.4.2).");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R369
            site.CaptureRequirement(
                    369,
                    @"[In FileNodeListFragment] rgFileNodes (variable): A stream of bytes that contains a sequence of FileNode structures (section 2.4.3).");
            
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R371
            site.CaptureRequirementIfIsTrue(
                    fileNodeListFragment.padding.Length <= 4,
                    371,
                    @"[In FileNodeListFragment] [rgFileNodes] [The stream is terminated when any of the following conditions is met:] The number of bytes between the end of the last read FileNode and the nextFragment field is less than 4 bytes.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R376
            site.CaptureRequirementIfIsInstanceOfType(
                    fileNodeListFragment.padding,
                    typeof(byte[]),
                    376,
                    @"[In FileNodeListFragment] padding (variable): An optional array of bytes between the last FileNode structure in the rgFileNodes field and the nextFragment field.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R378
            site.CaptureRequirementIfIsInstanceOfType(
                    fileNodeListFragment.nextFragment,
                    typeof(FileChunkReference64x32),
                    378,
                    @"[In FileNodeListFragment] nextFragment (12 bytes): A FileChunkReference64x32 structure (section 2.2.4.4) that specifies whether there are more fragments in this file node list, and if so, the location and size of the next fragment.");

            this.VerifyFileChunkReference64x32(fileNodeListFragment.nextFragment, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R381
            site.CaptureRequirement(
                    381,
                    @"[In FileNodeListFragment] The location of the nextFragment field is calculated by adding the size of this FileNodeListFragment structure minus the size of the nextFragment and footer fields to the location of this FileNodeListFragment structure.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R382
            site.CaptureRequirementIfIsInstanceOfType(
                    fileNodeListFragment.footer,
                    typeof(ulong),
                    382,
                    @"[In FileNodeListFragment] footer (8 bytes): An unsigned integer;");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R383
            site.CaptureRequirementIfAreEqual<ulong>(
                    0x8BC215C38233BA4B,
                    fileNodeListFragment.footer,
                    383,
                    @"[In FileNodeListFragment] [footer] MUST be ""0x8BC215C38233BA4B"".");

            this.VerifyFileNodeListHeader(fileNodeListFragment.Header, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileNodeListFragment structure.
        /// </summary>
        /// <param name="header">The instance of FileNodeListHeader structure.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileNodeListHeader(FileNodeListHeader header, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R387
            site.CaptureRequirementIfIsTrue(
                    header.uintMagic.GetType()==typeof(ulong) && header.uintMagic== 0xA4567AB1F5F7F4C4,
                    387,
                    @"[In FileNodeListHeader] uintMagic (8 bytes): An unsigned integer; MUST be ""0xA4567AB1F5F7F4C4"".");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R388
            site.CaptureRequirementIfIsInstanceOfType(
                    header.FileNodeListID,
                    typeof(uint),
                    388,
                    @"[In FileNodeListHeader] FileNodeListID (4 bytes): An unsigned integer that specifies the identity of the file node list (section 2.4) this fragment belongs to.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R389
            site.CaptureRequirementIfIsTrue(
                    header.FileNodeListID>= 0x00000010,
                    389,
                    @"[In FileNodeListHeader] MUST be equal to or greater than 0x00000010.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R392
            site.CaptureRequirementIfIsInstanceOfType(
                    header.nFragmentSequence,
                    typeof(uint),
                    392,
                    @"[In FileNodeListHeader] nFragmentSequence (4 bytes): An unsigned integer that specifies the index of the fragment in the file node list containing the fragment.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileNode structure.
        /// </summary>
        /// <param name="fileNode">The instance of FileNode structure.</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileNode(FileNode fileNode,string extension,ITestSite site)
        {
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R396 
            site.CaptureRequirement(
                    396,
                    @"[In FileNode] A FileNode structure is divided into header fields and a data field, fnd.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R400
            site.CaptureRequirement(
                    400,
                    @"[In FileNode] FileNodeID (10 bits): An unsigned integer that specifies the type of this FileNode structure.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R402
            site.CaptureRequirement(
                    402,
                    @"[In FileNode] Size (13 bits): An unsigned integer that specifies the size, in bytes, of this FileNode structure.");

            #region verify StpFormat and CbFormat
            if(fileNode.BaseType==1 || fileNode.BaseType == 2)
            {
                if (fileNode.StpFormat == 0)
                {
                    // If the OneNote file parse successfully and the StpFormat is 0, this requirement will be verified.
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R407
                    site.CaptureRequirement(
                            407,
                            @"[In FileNode] Value 0 means 8 bytes, uncompressed.");
                }
                if (fileNode.StpFormat == 1)
                {
                    // If the OneNote file parse successfully and the StpFormat is 1, this requirement will be verified.
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R408
                    site.CaptureRequirement(
                            408,
                            @"[In FileNode] Value 1 means 4 bytes, uncompressed.");
                }
                if (fileNode.StpFormat == 2)
                {
                    // If the OneNote file parse successfully and the StpFormat is 2, this requirement will be verified.
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R409
                    site.CaptureRequirement(
                            409,
                            @"[In FileNode] Value 2 means 2 bytes, compressed.");
                }
                if (fileNode.StpFormat == 3)
                {
                    // If the OneNote file parse successfully and the StpFormat is 3, this requirement will be verified.
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R410
                    site.CaptureRequirement(
                            410,
                            @"[In FileNode] Value 3 means 4 bytes, compressed.");
                }

                if (fileNode.CbFormat == 0)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R416
                    site.CaptureRequirement(
                            416,
                            @"[In FileNode] Value 0 means 4 bytes, uncompressed.");
                }
                if (fileNode.CbFormat == 1)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R417
                    site.CaptureRequirement(
                            417,
                            @"[In FileNode] Value 1 means 8 bytes, uncompressed.");
                }
                if (fileNode.CbFormat == 2)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R418
                    site.CaptureRequirement(
                            418,
                            @"[In FileNode] Value 2 means 1 byte, compressed.");
                }
                if (fileNode.CbFormat == 3)
                {
                    // Verify MS-ONESTORE requirement: MS-ONESTORE_R419
                    site.CaptureRequirement(
                            419,
                            @"[In FileNode] Value 3 means 2 bytes, compressed.");
                }
            }
            #endregion

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R402
            site.CaptureRequirement(
                    422,
                    @"[In FileNode] C - BaseType (4 bits): An unsigned integer that specifies whether the structure specified by fnd contains a FileNodeChunkReference structure (section 2.2.4.2).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R423
            site.CaptureRequirementIfIsTrue(
                    fileNode.BaseType==0 || fileNode.BaseType == 1 || fileNode.BaseType == 2,
                    423,
                    @"[In FileNode] [C - BaseType] MUST be one of the values [0,1,2] described in the following table.");


            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R434
            site.CaptureRequirement(
                    434,
                    @"[In FileNode] fnd (variable): A field that specifies additional data for this FileNode structure, if present.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R435
            site.CaptureRequirement(
                    435,
                    @"[In FileNode] The type of structure is specified by the value of the FileNodeID field.");

            if(fileNode.FileNodeID== FileNodeIDValues.ObjectSpaceManifestRootFND)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R438
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType==0 && (extension==".one" || extension==".onetoc2"),
                        438,
                        @"[In FileNode] FileNodeID value 0x004 means basetype is 0, Fnd structure is ObjectSpaceManifestRootFND (section 2.5.1), allowed file format is one and onetoc2");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.ObjectSpaceManifestListReferenceFND)
            {
                this.VerifyObjectSpaceManifestListReferenceFND((ObjectSpaceManifestListReferenceFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R439
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType==2 && (extension == ".one" || extension == ".onetoc2"),
                        439,
                        @"[In FileNode] FileNodeID value 0x008 means basetype is 2, Fnd structure is ObjectSpaceManifestListReferenceFND (section 2.5.2), allowed file format is one and onetoc2");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectSpaceManifestListStartFND)
            {
                this.VerifyObjectSpaceManifestListStartFND((ObjectSpaceManifestListStartFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R440
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType==0 && (extension == ".one" || extension == ".onetoc2"),
                        440,
                        @"[In FileNode] FileNodeID value 0x00C means basetype is 0, Fnd structure is ObjectSpaceManifestListStartFND (section 2.5.3), allowed file format is one and onetoc2");
            }

            if(fileNode.FileNodeID == FileNodeIDValues.RevisionManifestListReferenceFND)
            {
                this.VerifyRevisionManifestListReferenceFND((RevisionManifestListReferenceFND)fileNode.fnd, site);
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R441
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 2 && (extension == ".one" || extension == ".onetoc2"),
                        441,
                        @"[In FileNode] FileNodeID value 0x010 means basetype is 2, Fnd structure is RevisionManifestListReferenceFND (section 2.5.4), allowed file format is one and onetoc2");
            }

            if(fileNode.FileNodeID == FileNodeIDValues.RevisionManifestListStartFND)
            {
                this.VerifyRevisionManifestListStartFND((RevisionManifestListStartFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R442
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension == ".onetoc2"),
                        442,
                        @"[In FileNode] FileNodeID value 0x014 means basetype is 0, Fnd structure is RevisionManifestListStartFND (section 2.5.5), allowed file format is one and onetoc2.");
            }

            if(fileNode.FileNodeID== FileNodeIDValues.RevisionManifestStart4FND)
            {
                this.VerifyRevisionManifestStart4FND((RevisionManifestStart4FND)fileNode.fnd, site);
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R443
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && extension == ".onetoc2",
                        443,
                        @"[In FileNode] FileNodeID value 0x01B means basetype is 0, Fnd structure is RevisionManifestStart4FND (section 2.5.6), allowed file format is onetoc2.");
            }

            if(fileNode.FileNodeID== FileNodeIDValues.RevisionManifestEndFND)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R444
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension == ".onetoc2"),
                        444,
                        @"[In FileNode] FileNodeID value 0x01C means basetype is 0, Fnd structure is RevisionManifestEndFND, allowed file format is one and onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R446
                site.CaptureRequirementIfIsNull(
                        fileNode.fnd,
                        446,
                        @"[In FileNode] MUST contain no data.");
            }
           
            if(fileNode.FileNodeID== FileNodeIDValues.RevisionManifestStart6FND)
            {
                this.VerifyRevisionManifestStart6FND((RevisionManifestStart6FND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R447
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && extension == ".one",
                        447,
                        @"[In FileNode] FileNodeID value 0x01E means basetype is 0, Fnd structure is RevisionManifestStart6FND (section 2.5.7) allowed file format is one.");
            }

            if(fileNode.FileNodeID == FileNodeIDValues.RevisionManifestStart7FND)
            {
                this.VerifyRevisionManifestStart7FND((RevisionManifestStart7FND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R448
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && extension == ".one",
                        448,
                        @"[In FileNode] FileNodeID value 0x01F means basetype is 0, Fnd structure is RevisionManifestStart7FND (section 2.5.8), allowed file format is one.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.GlobalIdTableStartFNDX)
            {
                this.VerifyGlobalIdTableStartFNDX((GlobalIdTableStartFNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R449
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && extension == ".onetoc2",
                        449,
                        @"[In FileNode] FileNodeID value 0x021 means basetypeis 0, Fnd structure is GlobalIdTableStartFNDX (section 2.5.9), allowed file format is onetoc2");
            }
            if(fileNode.FileNodeID == FileNodeIDValues.GlobalIdTableStart2FND)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R450
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && extension == ".one",
                        450,
                        @"[In FileNode] FileNodeID value 0x022 means basetype is 0, Fnd structure is GlobalIdTableStart2FND allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R452
                site.CaptureRequirementIfIsNull(
                        fileNode.fnd,
                        452,
                        @"[In FileNode]  [FileNodeID value 0x022] MUST contain no data.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.GlobalIdTableEntryFNDX)
            {
                this.VerifyGlobalIdTableEntryFNDX((GlobalIdTableEntryFNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R453
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension==".onetoc2"),
                        453,
                        @"[In FileNode] FileNodeID value 0x024 means basetype is 0, Fnd structure is GlobalIdTableEntryFNDX (section 2.5.10), allowed file format is one and onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R545
                site.CaptureRequirementIfAreEqual<uint>(
                        0x024,
                        (uint)fileNode.FileNodeID,
                        545,
                        @"[In GlobalIdTableEntryFNDX] The value of the FileNode.FileNodeID field MUST be 0x024.");

            }
            if(fileNode.FileNodeID== FileNodeIDValues.GlobalIdTableEntry2FNDX)
            {
                this.VerifyGlobalIdTableEntry2FNDX((GlobalIdTableEntry2FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R454
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".onetoc2"),
                        454,
                        @"[In FileNode] FileNodeID value 0x025 means basetype is 0, Fnd structure is GlobalIdTableEntry2FNDX (section 2.5.11), allowed file format is onetoc2.");
               
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R553
                site.CaptureRequirementIfAreEqual<uint>(
                        0x025,
                        (uint)fileNode.FileNodeID,
                        553,
                        @"[In GlobalIdTableEntry2FNDX] The value of the FileNode.FileNodeID field MUST be 0x025.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.GlobalIdTableEntry3FNDX)
            {
                this.VerifyGlobalIdTableEntry3FNDX((GlobalIdTableEntry3FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R455
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".onetoc2"),
                        455,
                        @"[In FileNode] FileNodeID value 0x026 means basetype is 0, Fnd structure is GlobalIdTableEntry3FNDX (section 2.5.12), allowed file format is onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R562
                site.CaptureRequirementIfAreEqual<uint>(
                        0x026,
                        (uint)fileNode.FileNodeID,
                        562,
                        @"[In GlobalIdTableEntry3FNDX] The value of the FileNode.FileNodeID field MUST be 0x026.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.GlobalIdTableEndFNDX)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R456
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension == ".onetoc2"),
                        456,
                        @"[In FileNode] FileNodeID value 0x028 means basetype is 0, Fnd structure is GlobalIdTableEndFNDX, allowed file format is one and onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R458
                site.CaptureRequirementIfIsNull(
                        fileNode.fnd,
                        458,
                        @"[In FileNode] MUST contain no data.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectDeclarationWithRefCountFNDX)
            {
                this.VerifyObjectDeclarationWithRefCountFNDX((ObjectDeclarationWithRefCountFNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R459
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".onetoc2"),
                        459,
                        @"[In FileNode] FileNodeID value 0x02D means basetype is 1, Fnd structure is ObjectDeclarationWithRefCountFNDX (section 2.5.23), allowed file format is onetoc2.");
                
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R639
                site.CaptureRequirementIfAreEqual<uint>(
                        0x02D,
                        (uint)fileNode.FileNodeID,
                        639,
                        @"[In ObjectDeclarationWithRefCountFNDX] The value of the FileNode.FileNodeID field MUST be 0x02D.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.ObjectDeclarationWithRefCount2FNDX)
            {
                this.VerifyObjectDeclarationWithRefCount2FNDX((ObjectDeclarationWithRefCount2FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R460
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".onetoc2"),
                        460,
                        @"[In FileNode] FileNodeID value 0x02E means basetype is 1, Fnd structure is ObjectDeclarationWithRefCount2FNDX (section 2.5.24), allowed file format is onetoc2.");
                
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R645
                site.CaptureRequirementIfAreEqual<uint>(
                        0x02E,
                        (uint)fileNode.FileNodeID,
                        645,
                        @"[In ObjectDeclarationWithRefCount2FNDX] The value of the FileNode.FileNodeID field MUST be 0x02E.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCountFNDX)
            {
                this.VerifyObjectRevisionWithRefCountFNDX((ObjectRevisionWithRefCountFNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R461
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".onetoc2"),
                        461,
                        @"[In FileNode] FileNodeID value 0x041 means basetype is 1, Fnd structure is ObjectRevisionWithRefCountFNDX (section 2.5.13), allowed file format is onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R572
                site.CaptureRequirementIfAreEqual<uint>(
                        0x041,
                        (uint)fileNode.FileNodeID,
                        572,
                        @"[In ObjectRevisionWithRefCountFNDX]  The value of the FileNode.FileNodeID field MUST be 0x041. ");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectRevisionWithRefCount2FNDX)
            {
                this.VerifyObjectRevisionWithRefCount2FNDX((ObjectRevisionWithRefCount2FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R462
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".onetoc2"),
                        462,
                        @"[In FileNode] FileNodeID value 0x042 means basetype is 1, Fnd structure is ObjectRevisionWithRefCount2FNDX (section 2.5.14), allowed file format is onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R581
                site.CaptureRequirementIfAreEqual<uint>(
                        0x042,
                        (uint)fileNode.FileNodeID,
                        581,
                        @"[In ObjectRevisionWithRefCount2FNDX] The value of the FileNode.FileNodeID field MUST be 0x042.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.RootObjectReference2FNDX)
            {
                this.VerifyRootObjectReference2FNDX((RootObjectReference2FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R463
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".onetoc2"),
                        463,
                        @"[In FileNode] FileNodeID value 0x059 means basetype is 0, Fnd structure is RootObjectReference2FNDX (section 2.5.15), allowed file format is onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R591
                site.CaptureRequirementIfAreEqual<uint>(
                        0x059,
                        (uint)fileNode.FileNodeID,
                        591,
                        @"[In RootObjectReference2FNDX] The value of the FileNode.FileNodeID field MUST be 0x059.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.RootObjectReference3FND)
            {
                this.VerifyRootObjectReference3FND((RootObjectReference3FND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R464
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        464,
                        @"[In FileNode] FileNodeID value 0x05A means basetype is 0, Fnd structure is RootObjectReference3FND (section 2.5.16), allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R596
                site.CaptureRequirementIfAreEqual<uint>(
                        0x05A,
                        (uint)fileNode.FileNodeID,
                        596,
                        @"[In RootObjectReference3FND] The value of the FileNode.FileNodeID field MUST be 0x05A.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.RevisionRoleDeclarationFND)
            {
                this.VerifyRevisionRoleDeclarationFND((RevisionRoleDeclarationFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R465
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension == ".onetoc2"),
                        465,
                        @"[In FileNode] FileNodeID value 0x05C means basetype is 0, Fnd structure is RevisionRoleDeclarationFND (section 2.5.17), allowed file format is one and onetoc2.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.RevisionRoleAndContextDeclarationFND)
            {
                this.VerifyRevisionRoleAndContextDeclarationFND((RevisionRoleAndContextDeclarationFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R466
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        466,
                        @"[In FileNode] FileNodeID value 0x05D means basetype is 0, Fnd structure is RevisionRoleAndContextDeclarationFND (section 2.5.18), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3RefCountFND)
            {
                this.VerifyObjectDeclarationFileData3RefCountFND((ObjectDeclarationFileData3RefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R467
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        467,
                        @"[In FileNode] FileNodeID value 0x072 means basetype is 0, Fnd structure is ObjectDeclarationFileData3RefCountFND (section 2.5.27), allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R664
                site.CaptureRequirementIfAreEqual<uint>(
                        0x072,
                        (uint)fileNode.FileNodeID,
                        664,
                        @"[In ObjectDeclarationFileData3RefCountFND] The value of the FileNode.FileNodeID field MUST be 0x072. This structure has the following format.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectDeclarationFileData3LargeRefCountFND)
            {
                this.VerifyObjectDeclarationFileData3LargeRefCountFND((ObjectDeclarationFileData3LargeRefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R468
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        468,
                        @"[In FileNode] FileNodeID value 0x073 means basetype is 0, Fnd structure is ObjectDeclarationFileData3LargeRefCountFND (section 2.5.28), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectDataEncryptionKeyV2FNDX)
            {
                this.VerifyObjectDataEncryptionKeyV2FNDX((ObjectDataEncryptionKeyV2FNDX)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R469
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        469,
                        @"[In FileNode] FileNodeID value 0x07C means basetype is 1, Fnd structure is ObjectDataEncryptionKeyV2FNDX (section 2.5.19), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectInfoDependencyOverridesFND)
            {
                this.VerifyObjectInfoDependencyOverridesFND((ObjectInfoDependencyOverridesFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R470
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one" || extension == ".onetoc2"),
                        470,
                        @"[In FileNode] FileNodeID value 0x084 means basetype is 1, Fnd structure is ObjectInfoDependencyOverridesFND (section 2.5.20), allowed file format is one and onetoc2.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.DataSignatureGroupDefinitionFND)
            {
                this.VerifyDataSignatureGroupDefinitionFND((DataSignatureGroupDefinitionFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R471
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one" || extension == ".onetoc2"),
                        471,
                        @"[In FileNode] FileNodeID value 0x08C means basetype is 0, Fnd structure is DataSignatureGroupDefinitionFND (section 2.5.33), allowed file format is one and onetoc2.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.FileDataStoreListReferenceFND)
            {
                this.VerifyFileDataStoreListReferenceFND((FileDataStoreListReferenceFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R472
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 2 && (extension == ".one"),
                        472,
                        @"[In FileNode] FileNodeID value 0x090 means basetype is 2, Fnd structure is FileDataStoreListReferenceFND (section 2.5.21), allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R629
                site.CaptureRequirementIfAreEqual<uint>(
                        0x090,
                        (uint)fileNode.FileNodeID,
                        629,
                        @"[In FileDataStoreListReferenceFND] The value of the FileNode.FileNodeID field MUST be 0x090.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.FileDataStoreObjectReferenceFND)
            {
                this.VerifyFileDataStoreObjectReferenceFND((FileDataStoreObjectReferenceFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R473
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        473,
                        @"[In FileNode] FileNodeID value 0x094 means basetype is 1, Fnd structure is FileDataStoreObjectReferenceFND (section 2.5.22), allowed file format is one.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.ObjectDeclaration2RefCountFND)
            {
                this.VerifyObjectDeclaration2RefCountFND((ObjectDeclaration2RefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R474
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        474,
                        @"[In FileNode] FileNodeID value 0x0A4 means basetype is 1, Fnd structure is ObjectDeclaration2RefCountFND (section 2.5.25), allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R651
                site.CaptureRequirementIfAreEqual<uint>(
                        0x0A4,
                        (uint)fileNode.FileNodeID,
                        651,
                        @"[In ObjectDeclaration2RefCountFND] The value of the FileNode.FileNodeID field MUST be 0x0A4.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectDeclaration2LargeRefCountFND)
            {
                this.VerifyObjectDeclaration2LargeRefCountFND((ObjectDeclaration2LargeRefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R475
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        475,
                        @"[In FileNode] FileNodeID value 0x0A5 means basetype is 1, Fnd structure is ObjectDeclaration2LargeRefCountFND (section 2.5.26), allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R657
                site.CaptureRequirementIfAreEqual<uint>(
                        0x0A5,
                        (uint)fileNode.FileNodeID,
                        657,
                        @"[In ObjectDeclaration2LargeRefCountFND] The value of the FileNode.FileNodeID field MUST be 0x0A5.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectGroupListReferenceFND)
            {
                this.VerifyObjectGroupListReferenceFND((ObjectGroupListReferenceFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R476
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 2 && (extension == ".one"),
                        476,
                        @"[In FileNode] FileNodeID value 0x0B0 means basetype is 2, Fnd structure is ObjectGroupListReferenceFND (section 2.5.31), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectGroupStartFND)
            {
                this.VerifyObjectGroupStartFND((ObjectGroupStartFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R477
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        477,
                        @"[In FileNode] FileNodeID value 0x0B4 means basetype is 0, Fnd structure is ObjectGroupStartFND (section 2.5.32), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ObjectGroupEndFND)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R478
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 0 && (extension == ".one"),
                        478,
                        @"[In FileNode] FileNodeID value 0x0B8 means basetype is 0, Fnd structure is ObjectGroupEndFND, allowed file format is one.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R480
                site.CaptureRequirementIfIsNull(
                        fileNode.fnd,
                        480,
                        @"[In FileNode]  [FileNodeID value 0x0B8] MUST contain no data.");
            }
            if(fileNode.FileNodeID== FileNodeIDValues.HashedChunkDescriptor2FND)
            {
                this.VerifyHashedChunkDescriptor2FND((HashedChunkDescriptor2FND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R481
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        481,
                        @"[In FileNode] FileNodeID value 0x0C2 means basetype is 1, Fnd structure is HashedChunkDescriptor2FND (section 2.3.4.1), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2RefCountFND)
            {
                this.VerifyReadOnlyObjectDeclaration2RefCountFND((ReadOnlyObjectDeclaration2RefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R482
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        482,
                        @"[In FileNode] FileNodeID value 0x0C4 means basetype is 1, Fnd structure is ReadOnlyObjectDeclaration2RefCountFND (section 2.5.29), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ReadOnlyObjectDeclaration2LargeRefCountFND)
            {
                this.VerifyReadOnlyObjectDeclaration2LargeRefCountFND((ReadOnlyObjectDeclaration2LargeRefCountFND)fileNode.fnd, site);

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R483
                site.CaptureRequirementIfIsTrue(
                        fileNode.BaseType == 1 && (extension == ".one"),
                        483,
                        @"[In FileNode] FileNodeID value 0x0C5 means basetype is 1, Fnd structure is ReadOnlyObjectDeclaration2LargeRefCountFND (section 2.5.30), allowed file format is one.");
            }
            if (fileNode.FileNodeID == FileNodeIDValues.ChunkTerminatorFND)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R484
                site.CaptureRequirementIfIsTrue(
                        extension == ".one" || extension == ".onetoc2",
                        484,
                        @"[In FileNode] FileNodeID value 0x0FF means Fnd structure is ChunkTerminatorFND, allowed file format is one and onetoc2.");

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R486
                site.CaptureRequirementIfIsNull(
                        fileNode.fnd,
                        486,
                        @"[In FileNode] [FileNodeID value 0x0FF] MUST contain no data.");

            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectSpaceManifestRootFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectSpaceManifestRootFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectSpaceManifestRootFND(ObjectSpaceManifestRootFND fnd,ObjectSpaceManifestListReferenceFND objectSpaceManifestListReferenceFND, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R492
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gosidRoot,
                    typeof(ExtendedGUID),
                    492,
                    @"[In ObjectSpaceManifestRootFND] gosidRoot (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of the root object space.");

            this.VerifyExtendedGUID(fnd.gosidRoot, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R492
            site.CaptureRequirementIfAreEqual<ExtendedGUID>(
                    objectSpaceManifestListReferenceFND.gosid,
                    fnd.gosidRoot,
                    493,
                    @"[In ObjectSpaceManifestRootFND] This value MUST be equal to the ObjectSpaceManifestListReferenceFND.gosid field (section 2.5.2) of an object space within the object space manifest list (section 2.1.6).");

            this.VerifyExtendedGUID(objectSpaceManifestListReferenceFND.gosid, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectSpaceManifestListReferenceFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectSpaceManifestListReferenceFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectSpaceManifestListReferenceFND(ObjectSpaceManifestListReferenceFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R496
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.refField,
                    typeof(FileNodeChunkReference),
                    496,
                    @"[In ObjectSpaceManifestListReferenceFND] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies the location and size of the first FileNodeListFragment structure (section 2.4.1) in the object space manifest list.");

            this.VerifyFileNodeChunkReference(fnd.refField, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R497
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gosid,
                    typeof(ExtendedGUID),
                    497,
                    @"[In ObjectSpaceManifestListReferenceFND] gosid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of the object space (section 2.1.4) specified by the object space manifest list.");

            this.VerifyExtendedGUID(fnd.gosid, site);

            ExtendedGUID zeroExtendGuid = new ExtendedGUID();
            zeroExtendGuid.Guid = Guid.Empty;
            zeroExtendGuid.N = 0;

            //  Verify MS-ONESTORE requirement: MS-ONESTORE_R498
            site.CaptureRequirementIfAreNotEqual<ExtendedGUID>(
                    zeroExtendGuid,
                    fnd.gosid,
                    498,
                    @"[In ObjectSpaceManifestListReferenceFND] [gosid] MUST NOT be {{00000000-0000-0000-0000-000000000000},0} and MUST be unique relative to the other ObjectSpaceManifestListReferenceFND.gosid fields in this file.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with ObjectSpaceManifestListStartFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectSpaceManifestListStartFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectSpaceManifestListStartFND(ObjectSpaceManifestListStartFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R501
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gosid,
                    typeof(ExtendedGUID),
                    501,
                    @"[In ObjectSpaceManifestListStartFND] gosid (20 bytes): An ExtendedGUID structure that specifies the identity of the object space (section 2.1.4) being specified by this object space manifest list. ");

            this.VerifyExtendedGUID(fnd.gosid, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionManifestListReferenceFND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionManifestListReferenceFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestListReferenceFND(RevisionManifestListReferenceFND fnd, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R505
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.refField,
                    typeof(FileNodeChunkReference),
                    505,
                    @"[In RevisionManifestListReferenceFND] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies the location and size of the first FileNodeListFragment structure (section 2.4.1) in the revision manifest list.");

            this.VerifyFileNodeChunkReference(fnd.refField, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionManifestListStartFND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionManifestListStartFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestListStartFND(RevisionManifestListStartFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R508
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gosid,
                    typeof(ExtendedGUID),
                    508,
                    @"[In RevisionManifestListStartFND] gosid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of the object space (section 2.1.4) being revised by the revisions (section 2.1.8) in this list.");

            this.VerifyExtendedGUID(fnd.gosid, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionManifestStart4FND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionManifestStart4FND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestStart4FND(RevisionManifestStart4FND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R513
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.rid,
                    typeof(ExtendedGUID),
                    513,
                    @"[In RevisionManifestStart4FND] rid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of this revision (section 2.1.8).");

            this.VerifyExtendedGUID(fnd.rid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R515
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.ridDependent,
                    typeof(ExtendedGUID),
                    515,
                    @"[In RevisionManifestStart4FND] ridDependent (20 bytes): An ExtendedGUID structure that specifies the identity of a dependency revision.");

            this.VerifyExtendedGUID(fnd.ridDependent, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R519
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.RevisionRole,
                    typeof(int),
                    519,
                    @"[In RevisionManifestStart4FND] RevisionRole (4 bytes): An integer that specifies the revision role (section 2.1.12) that labels this revision (section 2.1.8).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R520
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.odcsDefault,
                    typeof(ushort),
                    520,
                    @"[In RevisionManifestStart4FND] odcsDefault (2 bytes): An unsigned integer that specifies whether the data contained by this revision manifest is encrypted. ");
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionManifestStart6FND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionManifestStart6FND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestStart6FND(RevisionManifestStart6FND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R524
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.rid,
                    typeof(ExtendedGUID),
                    524,
                    @"[In RevisionManifestStart6FND] rid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of this revision (section 2.1.8).");

            this.VerifyExtendedGUID(fnd.rid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R526
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.ridDependent,
                    typeof(ExtendedGUID),
                    526,
                    @"[In RevisionManifestStart6FND] ridDependent (20 bytes): An ExtendedGUID structure that specifies the identity of a dependency revision.");

            this.VerifyExtendedGUID(fnd.ridDependent, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R529
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.RevisionRole,
                    typeof(int),
                    529,
                    @"[In RevisionManifestStart6FND] RevisionRole (4 bytes): An integer that specifies the revision role (section 2.1.12) that labels this revision (section 2.1.8).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R530
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.odcsDefault,
                    typeof(ushort),
                    530,
                    @"[In RevisionManifestStart6FND] odcsDefault (2 bytes): An unsigned integer that specifies whether the data contained by this revision manifest is encrypted.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R531
            site.CaptureRequirementIfIsTrue(
                    fnd.odcsDefault== 0x0000 || fnd.odcsDefault == 0x0002,
                    531,
                    @"[In RevisionManifestStart6FND] [odcsDefault] MUST be one of the values described in the following table[0x0000, 0x0002].");

        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionManifestStart7FND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionManifestStart7FND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionManifestStart7FND(RevisionManifestStart7FND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R538
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Base,
                    typeof(RevisionManifestStart6FND),
                    538,
                    @"[In RevisionManifestStart7FND] base (46 bytes): A RevisionManifestStart6FND structure (section 2.5.7) that specifies the identity and other attributes of this revision (section 2.1.8).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R538
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gctxid,
                    typeof(ExtendedGUID),
                    539,
                    @"[In RevisionManifestStart7FND] gctxid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the context that labels this revision (section 2.1.8).");

            this.VerifyExtendedGUID(fnd.gctxid, site);
            // If R539 is verified, then R130 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R130
            site.CaptureRequirement(
                    130,
                    @"[In Context] It [A Context] is specified by an ExtendedGUID (section 2.2.1).");
        }
        /// <summary>
        /// This method is used to verify the requirements related with GlobalIdTableStartFNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of GlobalIdTableStartFNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdTableStartFNDX(GlobalIdTableStartFNDX fnd,ITestSite site)
        {

        }
        /// <summary>
        /// This method is used to verify the requirements related with GlobalIdTableEntryFNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of GlobalIdTableEntryFNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdTableEntryFNDX(GlobalIdTableEntryFNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R547 
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.index,
                    typeof(uint),
                    547,
                    @"[In GlobalIdTableEntryFNDX] index (4 bytes): An unsigned integer that specifies the index of the entry.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R548
            site.CaptureRequirementIfIsTrue(
                    fnd.index < 0xFFFFFF,
                    548,
                    @"[In GlobalIdTableEntryFNDX] [index] MUST be less than 0xFFFFFF.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R550
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.guid,
                    typeof(Guid),
                    550,
                    @"[In GlobalIdTableEntryFNDX] guid (16 bytes): A GUID, as specified by [MS-DTYP].");
        }
        /// <summary>
        /// This method is used to verify the requirements related with GlobalIdTableEntry2FNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of GlobalIdTableEntry2FNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdTableEntry2FNDX(GlobalIdTableEntry2FNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R555
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.iIndexMapFrom,
                    typeof(uint),
                    555,
                    @"[In GlobalIdTableEntry2FNDX] iIndexMapFrom (4 bytes): An unsigned integer that specifies the index of the entry in the dependency revision’s global identification table that is used to define this entry.");

            // If the R555 is verified, the R556 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R556
            site.CaptureRequirement(
                    556,
                    @"[In GlobalIdTableEntry2FNDX] [iIndexMapFrom] The index MUST be present in the global identification table of the dependency revision (section 2.1.8).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R557
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.iIndexMapTo,
                    typeof(uint),
                    557,
                    @"[In GlobalIdTableEntry2FNDX] iIndexMapTo (4 bytes): An unsigned integer that specifies the index of the entry in the current global identification table.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R558
            site.CaptureRequirementIfIsTrue(
                    fnd.iIndexMapTo< 0xFFFFFF,
                    558,
                    @"[In GlobalIdTableEntry2FNDX] [iIndexMapTo] MUST be less than 0xFFFFFF. ");
        }
        /// <summary>
        /// This method is used to verify the requirements related with GlobalIdTableEntry3FNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of GlobalIdTableEntry3FNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyGlobalIdTableEntry3FNDX(GlobalIdTableEntry3FNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R564
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.iIndexCopyFromStart,
                    typeof(uint),
                    564,
                    @"[In GlobalIdTableEntry3FNDX] iIndexCopyFromStart (4 bytes): An unsigned integer that specifies the index of the first entry in the range of entries in global identification table of the dependency revision (section 2.1.8).");

            // If the R564 is verified, the R565 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R565
            site.CaptureRequirement(
                    565,
                    @"[In GlobalIdTableEntry3FNDX] [iIndexCopyFromStart] The index MUST be present in the global identification table of the dependency revision.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R565
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cEntriesToCopy,
                    typeof(uint),
                    566,
                    @"[In GlobalIdTableEntry3FNDX] cEntriesToCopy (4 bytes): An unsigned integer that specifies the number of entries in the range.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R568
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.iIndexCopyToStart,
                    typeof(uint),
                    568,
                    @"[In GlobalIdTableEntry3FNDX] iIndexCopyToStart (4 bytes): An unsigned integer that specifies the index of the first entry in the range in the current global identification table.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclarationWithRefCountFNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclarationWithRefCountFNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclarationWithRefCountFNDX(ObjectDeclarationWithRefCountFNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R641
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.ObjectRef,
                    typeof(FileNodeChunkReference),
                    641,
                    @"[In ObjectDeclarationWithRefCountFNDX] ObjectRef (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1).");

            this.VerifyFileNodeChunkReference(fnd.ObjectRef, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R642
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.body,
                    typeof(ObjectDeclarationWithRefCountBody),
                    642,
                    @"[In ObjectDeclarationWithRefCountFNDX] body (10 bytes): An ObjectDeclarationWithRefCountBody structure (section 2.6.15) that specifies the identity and other attributes of this object.");

            this.VerifyObjectDeclarationWithRefCountBody(fnd.body, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R643
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(byte),
                    643,
                    @"[In ObjectDeclarationWithRefCountFNDX] cRef (1 byte): An unsigned integer that specifies the number of objects that reference this object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclarationWithRefCount2FNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclarationWithRefCount2FNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclarationWithRefCount2FNDX(ObjectDeclarationWithRefCount2FNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R647
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.ObjectRef,
                    typeof(FileNodeChunkReference),
                    647,
                    @"[In ObjectDeclarationWithRefCount2FNDX] ObjectRef (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1).");

            this.VerifyFileNodeChunkReference(fnd.ObjectRef, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R648
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.body,
                    typeof(ObjectDeclarationWithRefCountBody),
                    648,
                    @"[In ObjectDeclarationWithRefCount2FNDX] body (10 bytes): An ObjectDeclarationWithRefCountBody structure (section 2.6.15) that specifies the identity and other attributes of this object (section 2.1.5).");

            this.VerifyObjectDeclarationWithRefCountBody(fnd.body, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R649
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(uint),
                    649,
                    @"[In ObjectDeclarationWithRefCount2FNDX] cRef (4 bytes): An unsigned integer that specifies the reference count for this object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectRevisionWithRefCountFNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectRevisionWithRefCountFNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectRevisionWithRefCountFNDX(ObjectRevisionWithRefCountFNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R575
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    575,
                    @"[In ObjectRevisionWithRefCountFNDX] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1) containing the revised data for the object referenced by the oid field.");

            this.VerifyFileNodeChunkReference(fnd.Ref, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R576
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oid,
                    typeof(CompactID),
                    576,
                    @"[In ObjectRevisionWithRefCountFNDX] oid (4 bytes): A CompactID structure (section 2.2.2) that specifies the object that has been revised.");

            this.VerifyCompactID(fnd.oid, site);

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R577
            site.CaptureRequirement(
                    577,
                    @"[In ObjectRevisionWithRefCountFNDX] A - fHasOidReferences (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure referenced by the ref field contains references to other objects.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R578
            site.CaptureRequirement(
                    578,
                    @"[In ObjectRevisionWithRefCountFNDX] B - fHasOsidReferences (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure referenced by the ref field contains references to object spaces (section 2.1.4).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R579
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(uint),
                    579,
                    @"[In ObjectRevisionWithRefCountFNDX] cRef (6 bits): An unsigned integer that specifies the reference count for this object. ");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectRevisionWithRefCountFNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectRevisionWithRefCountFNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectRevisionWithRefCount2FNDX(ObjectRevisionWithRefCount2FNDX fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R584
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    584,
                    @"[In ObjectRevisionWithRefCount2FNDX] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1) containing the revised data for the object referenced by the oid field.");

            this.VerifyFileNodeChunkReference(fnd.Ref, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R585
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oid,
                    typeof(CompactID),
                    585,
                    @"[In ObjectRevisionWithRefCount2FNDX] oid (4 bytes): A CompactID structure (section 2.2.2) that specifies the object that has been revised.");

            this.VerifyCompactID(fnd.oid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R586
            site.CaptureRequirementIfIsTrue(
                    fnd.fHasOidReferences==0 || fnd.fHasOidReferences == 1,
                    586,
                    @"[In ObjectRevisionWithRefCount2FNDX] A - fHasOidReferences (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure referenced by the ref field contains references to other objects.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R587
            site.CaptureRequirementIfIsTrue(
                    fnd.fHasOsidReferences == 0 || fnd.fHasOsidReferences == 1,
                    587,
                    @"[In ObjectRevisionWithRefCount2FNDX] B - fHasOsidReferences (1 bit): A bit that specifies whether the ObjectSpaceObjectPropSet structure referenced by the ref field contains references to object spaces (section 2.1.4).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R588
            site.CaptureRequirementIfAreEqual<int>(
                    0,
                    fnd.Reserved,
                    588,
                    @"[In ObjectRevisionWithRefCount2FNDX] Reserved (30 bits): MUST be zero, and MUST be ignored.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R589
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(uint),
                    589,
                    @"[In ObjectRevisionWithRefCount2FNDX] cRef (4 bytes): An unsigned integer that specifies the reference count for this object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with RootObjectReference2FNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of RootObjectReference2FNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRootObjectReference2FNDX(RootObjectReference2FNDX fnd, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R593
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oidRoot,
                    typeof(CompactID),
                    593,
                    @"[In RootObjectReference2FNDX] oidRoot (4 bytes): A CompactID structure (section 2.2.2) that specifies the identity of the root object of the containing revision for the role specified by the RootRole field.");

            this.VerifyCompactID(fnd.oidRoot, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R594
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.RootRole,
                    typeof(uint),
                    594,
                    @"[In RootObjectReference2FNDX] RootRole (4 bytes): An unsigned integer that specifies the role of the root object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with RootObjectReference3FND structure.
        /// </summary>
        /// <param name="fnd">The instance of RootObjectReference3FND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRootObjectReference3FND(RootObjectReference3FND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R598
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oidRoot,
                    typeof(ExtendedGUID),
                    598,
                    @"[In RootObjectReference3FND] oidRoot (20 bytes): An ExtendedGUID (section 2.2.1) that specifies the identity of the root object of the containing revision for the role specified by the RootRole field.");

            this.VerifyExtendedGUID(fnd.oidRoot, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R599
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.RootRole,
                    typeof(uint),
                    599,
                    @"[In RootObjectReference3FND] RootRole (4 bytes): An unsigned integer that specifies the role of the root object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionRoleDeclarationFND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionRoleDeclarationFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionRoleDeclarationFND(RevisionRoleDeclarationFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R603
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.rid,
                    typeof(ExtendedGUID),
                    603,
                    @"[In RevisionRoleDeclarationFND] rid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of the revision to add the revision role to.");

            this.VerifyExtendedGUID(fnd.rid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R606
            site.CaptureRequirementIfIsTrue(
                    fnd.RevisionRole.GetType() == typeof(byte[]) && fnd.RevisionRole.Length == 4,
                    606,
                    @"[In RevisionRoleDeclarationFND] RevisionRole (4 bytes): Specifies a revision role for the default context.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R141
            site.CaptureRequirementIfIsTrue(
                    fnd.RevisionRole[3] == 0 && fnd.RevisionRole[2] == 0,
                    141,
                    @"[In Revision Role] It[A revision role] is specified by a 4-byte integer where the high 2 bytes MUST be set to zero.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with RevisionRoleAndContextDeclarationFND structure.
        /// </summary>
        /// <param name="fnd">The instance of RevisionRoleAndContextDeclarationFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyRevisionRoleAndContextDeclarationFND(RevisionRoleAndContextDeclarationFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R609
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Base,
                    typeof(RevisionRoleDeclarationFND),
                    609,
                    @"[In RevisionRoleAndContextDeclarationFND] base (24 bytes): A RevisionRoleDeclarationFND structure (section 2.5.17) that specifies the revision and revision role.");


            // Verify MS-ONESTORE requirement: MS-ONESTORE_R610
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.gctxid,
                    typeof(ExtendedGUID),
                    610,
                    @"[In RevisionRoleAndContextDeclarationFND] gctxid (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the context.");

            this.VerifyExtendedGUID(fnd.gctxid, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclarationFileData3RefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclarationFileData3RefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclarationFileData3RefCountFND(ObjectDeclarationFileData3RefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R666
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oid,
                    typeof(CompactID),
                    666,
                    @"[In ObjectDeclarationFileData3RefCountFND] oid (4 bytes): A CompactID structure (section 2.2.2) that specifies the identity of this object.");

            this.VerifyCompactID(fnd.oid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R667
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.jcid,
                    typeof(JCID),
                    667,
                    @"[In ObjectDeclarationFileData3RefCountFND] jcid (4 bytes): A JCID structure (section 2.6.14) that specifies the type of this object and the type of the data the object contains.");

            this.VerifyJCID(fnd.jcid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R668
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(byte),
                    668,
                    @"[In ObjectDeclarationFileData3RefCountFND] cRef (1 byte): An unsigned integer that specifies the reference count for this object (section 2.1.5).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R669
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.FileDataReference,
                    typeof(StringInStorageBuffer),
                    669,
                    @"[In ObjectDeclarationFileData3RefCountFND] FileDataReference (variable): A StringInStorageBuffer structure (section 2.2.3) that specifies the type and the target of the reference.");

            this.VerifyStringInStorageBuffer(fnd.FileDataReference, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R670
            site.CaptureRequirementIfIsTrue(
                    fnd.FileDataReference.StringData.StartsWith("<file>") || 
                    fnd.FileDataReference.StringData.StartsWith("<ifndf>") ||
                    fnd.FileDataReference.StringData.StartsWith("<invfdo>"),
                    670,
                    @"[In ObjectDeclarationFileData3RefCountFND] [FileDataReference] The value of the FileDataReference.StringData field MUST begin with one of the following strings: ""<file>""; ""<ifndf>""; ""<invfdo>"".");


        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclarationFileData3LargeRefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclarationFileData3LargeRefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclarationFileData3LargeRefCountFND(ObjectDeclarationFileData3LargeRefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R685
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oid,
                    typeof(CompactID),
                    685,
                    @"[In ObjectDeclarationFileData3LargeRefCountFND] oid (4 bytes): A CompactID structure (section 2.2.2) that specifies the identity of this object.");

            this.VerifyCompactID(fnd.oid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R686
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.jcid,
                    typeof(JCID),
                    686,
                    @"[In ObjectDeclarationFileData3LargeRefCountFND] jcid (4 bytes): A JCID structure (section 2.6.14) that specifies the type of this object and the type of the data the object contains.");

            this.VerifyJCID(fnd.jcid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R687
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(uint),
                    687,
                    @"[In ObjectDeclarationFileData3LargeRefCountFND] cRef (4 bytes): An unsigned integer that specifies the reference count for this objects (section 2.1.5).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R688
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.FileDataReference,
                    typeof(StringInStorageBuffer),
                    688,
                    @"[In ObjectDeclarationFileData3LargeRefCountFND] FileDataReference (variable): A StringInStorageBuffer structure (section 2.2.3) that specifies the type and the target of the reference.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R689
            site.CaptureRequirementIfIsTrue(
                    fnd.FileDataReference.StringData.StartsWith("<file>") ||
                    fnd.FileDataReference.StringData.StartsWith("<ifndf>") ||
                    fnd.FileDataReference.StringData.StartsWith("<invfdo>"),
                    689,
                    @"[In ObjectDeclarationFileData3LargeRefCountFND] [FileDataReference] The value of the FileDataReference.StringData field MUST begin with one of the following strings: [""<file>""; ""<ifndf>""; ""<invfdo>""].");


        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDataEncryptionKeyV2FNDX structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDataEncryptionKeyV2FNDX structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDataEncryptionKeyV2FNDX(ObjectDataEncryptionKeyV2FNDX fnd,ITestSite site)
        {

        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectInfoDependencyOverridesFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectInfoDependencyOverridesFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectInfoDependencyOverridesFND(ObjectInfoDependencyOverridesFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R621
            site.CaptureRequirementIfIsTrue(
                    fnd.data.SerializeToByteList().Count < 1024,
                    621,
                    @"[In ObjectInfoDependencyOverridesFND] The total size of the data field, in bytes, MUST be less than 1024; ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R624
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    624,
                    @"[In ObjectInfoDependencyOverridesFND] ref (variable): A FileNodeChunkReference structure that specifies the location of an ObjectInfoDependencyOverrideData structure (section 2.6.10) if the value of the ref field is not ""fcrNil"".");

            this.VerifyFileNodeChunkReference(fnd.Ref, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R625
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.data,
                    typeof(ObjectInfoDependencyOverrideData),
                    625,
                    @"[In ObjectInfoDependencyOverridesFND] data (variable): An optional ObjectInfoDependencyOverrideData structure (section 2.6.10) that specifies the updated reference counts for objects (section 2.1.5).");

            if(fnd.Ref.IsfcrNil())
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R626
                site.CaptureRequirementIfIsNotNull(
                        fnd.data,
                        626,
                        @"[In ObjectInfoDependencyOverridesFND]  [data] MUST exist if the value of the ref field is ""fcrNil"".");

                bool isVaildfcrNil = false;

                foreach(byte b in fnd.Ref.Stp)
                {
                    if(b==byte.MaxValue)
                    {
                        isVaildfcrNil = true;
                    }
                    else
                    {
                        isVaildfcrNil = false;
                        break;
                    }
                }
                if (isVaildfcrNil == true)
                {
                    foreach (byte b in fnd.Ref.Cb)
                    {
                        if (b != byte.MinValue)
                        {
                            isVaildfcrNil = false;
                            break;
                        }
                    }
                }

                // Verify MS-ONESTORE requirement: MS-ONESTORE_R187
                site.CaptureRequirementIfIsTrue(
                        isVaildfcrNil,
                        187,
                        @"[In File Chunk Reference] Special values:
fcrNil: Specifies a file chunk reference where all bits of the stp field are set to 1, and all bits of the cb field are set to zero.");
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with DataSignatureGroupDefinitionFND structure.
        /// </summary>
        /// <param name="fnd">The instance of DataSignatureGroupDefinitionFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyDataSignatureGroupDefinitionFND(DataSignatureGroupDefinitionFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R727
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.DataSignatureGroup,
                    typeof(ExtendedGUID),
                    727,
                    @"[In DataSignatureGroupDefinitionFND] DataSignatureGroup (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the signature. ");

            this.VerifyExtendedGUID(fnd.DataSignatureGroup, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileDataStoreListReferenceFND structure.
        /// </summary>
        /// <param name="fnd">The instance of FileDataStoreListReferenceFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileDataStoreListReferenceFND(FileDataStoreListReferenceFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R631
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    631,
                    @"[In FileDataStoreListReferenceFND] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to a FileNodeListFragment structure (section 2.4.1).");

            this.VerifyFileNodeChunkReference(fnd.Ref, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileDataStoreObjectReferenceFND structure.
        /// </summary>
        /// <param name="fnd">The instance of FileDataStoreObjectReferenceFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileDataStoreObjectReferenceFND(FileDataStoreObjectReferenceFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R635
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    635,
                    @"[In FileDataStoreObjectReferenceFND] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to a FileDataStoreObject structure (section 2.6.13).");

            this.VerifyFileNodeChunkReference((FileNodeChunkReference)fnd.Ref, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R636
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.guidReference,
                    typeof(Guid),
                    636,
                    @"[In FileDataStoreObjectReferenceFND] guidReference (16 bytes): A GUID, as specified by [MS-DTYP], that specifies the identity of this file data object.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclaration2RefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclaration2RefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclaration2RefCountFND(ObjectDeclaration2RefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R653
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.BlobRef,
                    typeof(FileNodeChunkReference),
                    653,
                    @"[In ObjectDeclaration2RefCountFND] BlobRef (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1).");
            this.VerifyFileNodeChunkReference(fnd.BlobRef, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R654
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.body,
                    typeof(ObjectDeclaration2Body),
                    654,
                    @"[In ObjectDeclaration2RefCountFND] body (9 bytes): An ObjectDeclaration2Body structure (section 2.6.16) that specifies the identity and other attributes of this object.");

            this.VerifyObjectDeclaration2Body(fnd.body, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R655
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(byte),
                    655,
                    @"[In ObjectDeclaration2RefCountFND] cRef (1 byte): An unsigned integer that specifies the reference count for this object (section 2.1.5).");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclaration2LargeRefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectDeclaration2LargeRefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclaration2LargeRefCountFND(ObjectDeclaration2LargeRefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R659
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.BlobRef,
                    typeof(FileNodeChunkReference),
                    659,
                    @"[In ObjectDeclaration2LargeRefCountFND] BlobRef (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies a reference to an ObjectSpaceObjectPropSet structure (section 2.6.1).");

            this.VerifyFileNodeChunkReference(fnd.BlobRef, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R660
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.body,
                    typeof(ObjectDeclaration2Body),
                    660,
                    @"[In ObjectDeclaration2LargeRefCountFND] body (9 bytes): An ObjectDeclaration2Body structure (section 2.6.16) that specifies the identity and other attributes of this object.");

            this.VerifyObjectDeclaration2Body(fnd.body, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R661
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.cRef,
                    typeof(uint),
                    661,
                    @"[In ObjectDeclaration2LargeRefCountFND] cRef (4 bytes): An unsigned integer that specifies the reference count for this object (section 2.1.5).");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectGroupListReferenceFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectGroupListReferenceFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectGroupListReferenceFND(ObjectGroupListReferenceFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R718
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Ref,
                    typeof(FileNodeChunkReference),
                    718,
                    @"[In ObjectGroupListReferenceFND] ref (variable): A FileNodeChunkReference structure (section 2.2.4.2) that specifies the location and size of the first FileNodeListFragment structure (section 2.4.1) in the file node list (section 2.4) of the object group.");

            this.VerifyFileNodeChunkReference(fnd.Ref, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R719
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.ObjectGroupID,
                    typeof(ExtendedGUID),
                    719,
                    @"[In ObjectGroupListReferenceFND] ObjectGroupID (20 bytes): An ExtendedGUID structure (section 2.2.1) that specifies the identity of the object group that the ref field value points to.");

            this.VerifyExtendedGUID(fnd.ObjectGroupID, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with ObjectGroupStartFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ObjectGroupStartFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectGroupStartFND(ObjectGroupStartFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R722
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.oid,
                    typeof(ExtendedGUID),
                    722,
                    @"[In ObjectGroupStartFND] oid (20 bytes): An ExtendedGUID (section 2.2.1) that specifies the identity of the object group.");

            this.VerifyExtendedGUID(fnd.oid, site);
        }
        /// <summary>
        /// This method is used to verify the requirements related with HashedChunkDescriptor2FND structure.
        /// </summary>
        /// <param name="fnd">The instance of HashedChunkDescriptor2FND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyHashedChunkDescriptor2FND(HashedChunkDescriptor2FND fnd,ITestSite site)
        {

        }
        /// <summary>
        /// This method is used to verify the requirements related with ReadOnlyObjectDeclaration2RefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ReadOnlyObjectDeclaration2RefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyReadOnlyObjectDeclaration2RefCountFND(ReadOnlyObjectDeclaration2RefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R705
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Base,
                    typeof(ObjectDeclaration2RefCountFND),
                    705,
                    @"[In ReadOnlyObjectDeclaration2RefCountFND] base (variable): An ObjectDeclaration2RefCountFND structure (section 2.5.25) that specifies the identity and other attributes of this object (section 2.1.5). ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R706
            site.CaptureRequirementIfIsTrue(
                    fnd.Base.body.jcid.IsPropertySet==1 && fnd.Base.body.jcid.IsReadOnly==1,
                    706,
                    @"[In ReadOnlyObjectDeclaration2RefCountFND] The values of the base.body.jcid.IsPropertySet and base.body.jcid.IsReadOnly fields MUST be true.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R707
            site.CaptureRequirementIfIsTrue(
                    fnd.md5Hash.Length==16,
                    707,
                    @"[In ReadOnlyObjectDeclaration2RefCountFND] md5Hash (16 bytes): An unsigned integer that specifies an MD5 checksum, as specified in [RFC1321], of the data referenced by the base.BlobRef field. ");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ReadOnlyObjectDeclaration2LargeRefCountFND structure.
        /// </summary>
        /// <param name="fnd">The instance of ReadOnlyObjectDeclaration2LargeRefCountFND structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyReadOnlyObjectDeclaration2LargeRefCountFND(ReadOnlyObjectDeclaration2LargeRefCountFND fnd,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R712
            site.CaptureRequirementIfIsInstanceOfType(
                    fnd.Base,
                    typeof(ObjectDeclaration2LargeRefCountFND),
                    712,
                    @"[In ReadOnlyObjectDeclaration2LargeRefCountFND] base (variable): An ObjectDeclaration2LargeRefCountFND structure (section 2.5.26) that specifies the identity and other attributes of this object (section 2.1.5).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R713
            site.CaptureRequirementIfIsTrue(
                    fnd.Base.body.jcid.IsPropertySet==1 && fnd.Base.body.jcid.IsReadOnly==1,
                    713,
                    @"[In ReadOnlyObjectDeclaration2LargeRefCountFND] The values of the base.body.jcid.IsPropertySet and base.body.jcid.IsReadOnly fields MUST be true.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R714
            site.CaptureRequirementIfIsTrue(
                    fnd.md5Hash.Length==16,
                    714,
                    @"[In ReadOnlyObjectDeclaration2LargeRefCountFND] md5Hash (16 bytes): An unsigned integer that specifies an MD5 checksum, as specified in [RFC1321], of the data referenced by the base.BlobRef field.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclarationWithRefCountBody structure.
        /// </summary>
        /// <param name="body">The instance of ObjectDeclarationWithRefCountBody structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclarationWithRefCountBody(ObjectDeclarationWithRefCountBody body,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R867
            site.CaptureRequirementIfIsInstanceOfType(
                    body.oid,
                    typeof(CompactID),
                    867,
                    @"[In ObjectDeclarationWithRefCountBody] oid (4 bytes): A CompactID structure (section 2.2.2) that specifies the identity of this object.");

            this.VerifyCompactID(body.oid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R869
            site.CaptureRequirementIfAreEqual<uint>(
                    0x01,
                    body.jci,
                    869,
                    @"[In ObjectDeclarationWithRefCountBody] MUST be 0x01.");

            // If R869 is verufied then the jci is an unsigned integer, so R868 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R868
            site.CaptureRequirement(
                    868,
                    @"[In ObjectDeclarationWithRefCountBody] jci (10 bits): An unsigned integer that specifies the value of the JCID.index field of the object. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R871
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    body.odcs,
                    871,
                    @"[In ObjectDeclarationWithRefCountBody] MUST be zero.");

            // If R871 is verufied then the odcs is an unsigned integer, so R870 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R870
            site.CaptureRequirement(
                    870,
                    @"[In ObjectDeclarationWithRefCountBody] odcs (4 bits): An unsigned integer that specifies whether the data contained by this object is encrypted. ");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R872
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    body.fReserved1,
                    872,
                    @"[In ObjectDeclarationWithRefCountBody] A - fReserved1 (2 bits): An unsigned integer that MUST be zero, and MUST be ignored.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R873
            site.CaptureRequirementIfIsTrue(
                    body.fHasOidReferences == 0 || body.fHasOidReferences == 1,
                    873,
                    @"[In ObjectDeclarationWithRefCountBody] B - fHasOidReferences (1 bit): Specifies whether this object contains references to other objects.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R874
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    body.fHasOsidReferences,
                    874,
                    @"[In ObjectDeclarationWithRefCountBody] C - fHasOsidReferences (1 bit): Specifies whether this object contains references to object spaces (section 2.1.4).");

            // If R874 is verufied then the fHasOsidReferences is zero, so R875 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R875
            site.CaptureRequirement(
                    875,
                    @"[In ObjectDeclarationWithRefCountBody] MUST be zero.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R876
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    body.fReserved2,
                    876,
                    @"[In ObjectDeclarationWithRefCountBody] fReserved2 (30 bits): An unsigned integer that MUST be zero, and MUST be ignored.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with ObjectDeclaration2Body structure.
        /// </summary>
        /// <param name="body">The instance of ObjectDeclaration2Body structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyObjectDeclaration2Body(ObjectDeclaration2Body body, ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R878
            site.CaptureRequirementIfIsInstanceOfType(
                    body.oid,
                    typeof(CompactID),
                    878,
                    @"[In ObjectDeclaration2Body] oid (4 bytes): A CompactID (section 2.2.2) that specifies the identity of this object.");

            this.VerifyCompactID(body.oid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R879
            site.CaptureRequirementIfIsInstanceOfType(
                    body.jcid,
                    typeof(JCID),
                    879,
                    @"[In ObjectDeclaration2Body] jcid (4 bytes): A JCID (section 2.6.14) that specifies the type of data this object contains.");

            this.VerifyJCID(body.jcid, site);

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R880
            site.CaptureRequirementIfIsTrue(
                    body.fHasOidReferences == 0 || body.fHasOidReferences == 1,
                    880,
                    @"[In ObjectDeclaration2Body] A - fHasOidReferences (1 bit): A bit that specifies whether this object contains references to other objects.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R881
            site.CaptureRequirementIfIsTrue(
                    body.fHasOsidReferences == 0 || body.fHasOsidReferences == 1,
                    881,
                    @"[In ObjectDeclaration2Body] B - fHasOsidReferences (1 bit): A bit that specifies whether this object contains references to object spaces (section 2.1.4) or contexts (section 2.1.11).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R882
            site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    body.fReserved2,
                    882,
                    @"[In ObjectDeclaration2Body] fReserved2 (6 bits): MUST be zero, and MUST be ignored.");
        }

        /// <summary>
        /// This method is used to verify the requirements related with JCID structure.
        /// </summary>
        /// <param name="jcid">The instance of JCID structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyJCID(JCID jcid,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R858
            site.CaptureRequirementIfIsInstanceOfType(
                    jcid.Index,
                    typeof(int),
                    858,
                    @"[In JCID] index (2 bytes): An unsigned integer that specifies the type of object.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R859
            site.CaptureRequirementIfIsTrue(
                    jcid.IsBinary == 0 || jcid.IsBinary == 1,
                    859,
                    @"[In JCID] A - IsBinary (1 bit): Specifies whether the object contains encryption data transmitted over the File Synchronization via SOAP over HTTP Protocol, as specified in [MS-FSSHTTP].");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R860
            site.CaptureRequirementIfIsTrue(
                    jcid.IsPropertySet == 0 || jcid.IsPropertySet == 1,
                    860,
                    @"[In JCID] B - IsPropertySet (1 bit): Specifies whether the object contains a property set (section 2.1.1).");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R861
            site.CaptureRequirementIfIsTrue(
                    jcid.IsGraphNode==0 || jcid.IsGraphNode==1,
                    861,
                    @"[In JCID] C - IsGraphNode (1 bit): Undefined and MUST be ignored.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R862
            site.CaptureRequirementIfIsTrue(
                    jcid.IsFileData == 0 || jcid.IsFileData == 1,
                    862,
                    @"[In JCID] D - IsFileData (1 bit): Specifies whether the object is a file data object. ");

            if (jcid.IsFileData == 1)
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R863
                site.CaptureRequirementIfIsTrue(
                        jcid.IsBinary == 0 && jcid.IsPropertySet == 0 && jcid.IsGraphNode == 0 && jcid.IsReadOnly == 0,
                        863,
                        @"[In JCID] If the value of IsFileData is ""true"", then the values of the IsBinary, IsPropertySet, IsGraphNode, and IsReadOnly fields MUST all be false.");
            }

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R864
            site.CaptureRequirementIfIsTrue(
                    jcid.IsReadOnly == 0 || jcid.IsReadOnly == 1,
                    864,
                    @"[In JCID] E - IsReadOnly (1 bit): Specifies whether the object's data MUST NOT be changed when the object is revised.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R865
            site.CaptureRequirementIfAreEqual<int>(
                    0,
                    jcid.Reserved,
                    865,
                    @"[In JCID] Reserved (11 bits): MUST be zero, and MUST be ignored.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with StringInStorageBuffer structure.
        /// </summary>
        /// <param name="stringInStorageBuffer">The instance of StringInStorageBuffer structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyStringInStorageBuffer(StringInStorageBuffer stringInStorageBuffer,ITestSite site)
        {
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R177
            site.CaptureRequirement(
                    177,
                    @"[In StringInStorageBuffer] The StringInStorageBuffer structure is a variable-length Unicode string.");

            //  Verify MS-ONESTORE requirement: MS-ONESTORE_R179
            site.CaptureRequirementIfIsInstanceOfType(
                    stringInStorageBuffer.Cch,
                    typeof(uint),
                    179,
                    @"[In StringInStorageBuffer] cch (4 bytes): An unsigned integer that specifies the number of characters in the string.");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R180
            site.CaptureRequirement(
                    180,
                    @"[In StringInStorageBuffer] StringData (variable): An array of UTF-16 Unicode characters.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R181 
            site.CaptureRequirementIfAreEqual<uint>(
                    stringInStorageBuffer.Cch,
                    (uint)stringInStorageBuffer.StringData.Length,
                    181,
                    @"[In StringInStorageBuffer] The length of the array MUST be equal to the value specified by the cch field.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with ExtendedGUID structure.
        /// </summary>
        /// <param name="instance">The instance of ExtendedGUID structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyExtendedGUID(ExtendedGUID instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R165
            site.CaptureRequirementIfIsTrue(
                    instance.Guid.GetType()==typeof(Guid) && instance.N.GetType()==typeof(uint),
                    165,
                    @"[In ExtendedGUID] The ExtendedGUID structure is a combination of a GUID, as specified by [MS-DTYP], and an unsigned integer.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R168
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Guid,
                    typeof(Guid),
                    168,
                    @"[In ExtendedGUID] guid (16 bytes): Specifies a GUID, as specified by [MS-DTYP].");

            if(instance.Guid==Guid.Parse("{00000000-0000-0000-0000-000000000000}"))
            {
                // Verify MS-ONESTORE requirement: MS-ONESTORE_R169
                site.CaptureRequirementIfAreEqual<uint>(
                        0,
                        instance.N,
                        169,
                        @"[In ExtendedGUID] n (4 bytes): An unsigned integer that MUST be zero when the guid field value is {00000000-0000-0000-0000-000000000000}.");
            }
        }
        /// <summary>
        /// This method is used to verify the requirements related with CompactID structure.
        /// </summary>
        /// <param name="instance">The instance of CompactID structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyCompactID(CompactID instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R174
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.N,
                    typeof(uint),
                    174,
                    @"[In CompactID] n (8 bits): An unsigned integer that specifies the value of the ExtendedGUID.n field.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R175
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.GuidIndex,
                    typeof(uint),
                    175,
                    @"[In CompactID] guidIndex (24 bits): An unsigned integer that specifies the index in the global identification table. ");

            // If R174 and R175 are verified, the R171 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R171
            site.CaptureRequirement(
                    171,
                    @"[In CompactID] The CompactID structure is a combination of two unsigned integers.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileChunkReference32 structure.
        /// </summary>
        /// <param name="instance">The instance of FileChunkReference32 structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileChunkReference32(FileChunkReference32 instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R191
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Stp,
                    typeof(uint),
                    191,
                    @"[In FileChunkReference32] stp (4 bytes): An unsigned integer that specifies the location of the referenced data in the file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R192
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Cb,
                    typeof(uint),
                    192,
                    @"[In FileChunkReference32] cb (4 bytes): An unsigned integer that specifies the size, in bytes, of the referenced data.");

            // If R191 and R192 are verified, the R189 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R171
            site.CaptureRequirement(
                    189,
                    @"[In FileChunkReference32] A FileChunkReference32 structure is a file chunk reference (section 2.2.4) where both the stp field and the cb field are 4 bytes in size.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileNodeChunkReference structure.
        /// </summary>
        /// <param name="instance">The instance of FileNodeChunkReference structure</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileNodeChunkReference(FileNodeChunkReference instance, ITestSite site)
        {
            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R194
            site.CaptureRequirement(
                    194,
                    @"[In FileNodeChunkReference] The size of the file chunk reference (section 2.2.4) is specified by the FileNode.StpFormat and FileNode.CbFormat fields of the FileNode structure that contains the FileNodeChunkReference structure.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R197
            site.CaptureRequirementIfIsTrue(
                    instance.Stp.Length == 2 || instance.Stp.Length == 4 || instance.Stp.Length == 8,
                    197,
                    @"[In FileNodeChunkReference] stp (variable): An unsigned integer that specifies the location of the referenced data in the file. ");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R198
            site.CaptureRequirement(
                    198,
                    @"[In FileNodeChunkReference] The size and meaning of the stp field is specified by the value of the FileNode.StpFormat field.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R199
            site.CaptureRequirementIfIsTrue(
                    instance.Cb.Length == 8 || instance.Cb.Length == 4 ||
                    instance.Cb.Length == 2 || instance.Cb.Length == 1,
                    199,
                    @"[In FileNodeChunkReference] cb (variable): An unsigned integer that specifies the size, in bytes, of the data. ");

            // If the OneNote file parse successfully, this requirement will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R200
            site.CaptureRequirement(
                    200,
                    @"[In FileNodeChunkReference] The size and meaning of the cb field is specified by the value of FileNode.CbFormat field.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileChunkReference64 structure.
        /// </summary>
        /// <param name="instance">The instance of FileChunkReference64</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileChunkReference64(FileChunkReference64 instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R203
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Stp,
                    typeof(ulong),
                    203,
                    @"[In FileChunkReference64] stp (8 bytes): An unsigned integer that specifies the location of the referenced data in the file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R204
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Cb,
                    typeof(ulong),
                    204,
                    @"[In FileChunkReference64] cb (8 bytes): An unsigned integer that specifies the size, in bytes, of the referenced data.");

            // If R203 and R204 are verified,so R201 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R201
            site.CaptureRequirement(
                    201,
                    @"[In FileChunkReference64] A FileChunkReference64 structure is a file chunk reference (section 2.2.4) where both the stp field and the cb field are 8 bytes in size.");
        }
        /// <summary>
        /// This method is used to verify the requirements related with FileChunkReference64x32 structure.
        /// </summary>
        /// <param name="instance">The instance of FileChunkReference64x32</param>
        /// <param name="site">Instance of ITestSite</param>
        private void VerifyFileChunkReference64x32(FileChunkReference64x32 instance,ITestSite site)
        {
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R207
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Stp,
                    typeof(ulong),
                    207,
                    @"[In FileChunkReference64x32] stp (8 bytes): An unsigned integer that specifies the location of the referenced data in the file.");

            // Verify MS-ONESTORE requirement: MS-ONESTORE_R208
            site.CaptureRequirementIfIsInstanceOfType(
                    instance.Cb,
                    typeof(uint),
                    208,
                    @"[In FileChunkReference64x32] cb (4 bytes): An unsigned integer that specifies the size, in bytes, of the referenced data.");

            // If R207 and R208 are verified,so R205 will be verified.
            // Verify MS-ONESTORE requirement: MS-ONESTORE_R205
            site.CaptureRequirement(
                    205,
                    @"[In FileChunkReference64x32] A FileChunkReference64x32 structure is a file chunk reference (section 2.2.4) where the stp field is 8 bytes in size and the cb field is 4 bytes in size.");
        }
    }
}
