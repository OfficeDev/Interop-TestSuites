namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestSuites.Common;
    using System;
    using System.Collections.Generic;
    
    /// <summary>
    /// This class specifies changes sub-request. 
    /// </summary>
    public class PutChangesCellSubRequest : FsshttpbCellSubRequest
    {
        /// <summary>
        /// Initializes a new instance of the PutChangesCellSubRequest class
        /// </summary>
        /// <param name="subRequestID">Specify the sub request id</param>
        /// <param name="storageIndexExGuid">Specify the storage index ExGuid.</param>
        public PutChangesCellSubRequest(ulong subRequestID, ExGuid storageIndexExGuid)
        {
            this.RequestID = subRequestID;
            this.RequestType = Convert.ToUInt64(RequestTypes.PutChanges);
            this.StorageIndexExtendedGUID = storageIndexExGuid;
            this.ExpectedStorageIndexExtendedGUID = new ExGuid();
            this.ImplyNullExpectedIfNoMapping = 0;
            this.Partial = 0;
            this.PartialLast = 0;
            this.FavorCoherencyFailureOverNotFound = 1;
            this.AbortRemainingPutChangesOnFailure = 0;
            this.Reserved1Bit = 0;
            this.ReturnCompleteKnowledgeIfPossible = 1;
            this.LastWriterWinsOnNextChange = 0;
            this.Reserve1Byte = 0;

            List<byte> byteList = new List<byte>();
            byteList.AddRange(new byte[1]);
            this.ContentVersionCoherencyCheck = new BinaryItem(byteList);

            List<StringItem> Content = new List<StringItem>();
            string str1 = "str1";
            StringItem str1Item = new StringItem();
            str1Item.Count = new Compact64bitInt((ulong)str1.Length);
            str1Item.Content = str1;
            Content.Add(str1Item);
            string str2 = "str2";
            StringItem str2Item = new StringItem();
            str2Item.Count = new Compact64bitInt((ulong)str2.Length);
            str2Item.Content = str2;
            Content.Add(str2Item);

            this.AuthorLogins = new StringItemArray((ulong)Content.Count, Content);        

            this.IsAdditionalFlagsUsed = false;
            this.IsLockIdUsed = false;
            this.IsDiagnosticRequestOptionInputUsed = false;
        }

        /// <summary>
        /// Gets or sets Storage Index Extended GUID (variable): An extended GUID that specifies the storage index.
        /// </summary>
        public ExGuid Reserved { get; set; }

        /// <summary>
        /// Gets or sets Storage Index Extended GUID (variable): An extended GUID that specifies the storage index.
        /// </summary>
        public ExGuid StorageIndexExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets Expected Storage Index Extended GUID (variable): An extended GUID that specifies the expected storage index.
        /// </summary>
        public ExGuid ExpectedStorageIndexExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets I - Imply Null Expected if No Mapping (1 bit): A bit that specifies that the expected storage index is the basis for what the client believes is the current state of the storage index if set to 1, otherwise the expected storage index is not specified.
        /// </summary>
        public int ImplyNullExpectedIfNoMapping { get; set; }

        /// <summary>
        /// Gets or sets P - Partial (1 bit): A bit that specifies that this is a partial put changes and not the full changes.
        /// </summary>
        public int Partial { get; set; }

        /// <summary>
        /// Gets or sets L - Partial Last (1 bit): A bit that specifies if this is the last put changes in a partial set of changes.
        /// </summary>
        public int PartialLast { get; set; }

        /// <summary>
        /// Gets or sets F - Favor Coherency Failure Over Not Found (1 bit): A bit that specifies to force a coherency check on the server if a Referenced Data Element Not Found. This may result in a Coherency Failure returned instead of Referenced Data Element Not Found. 
        /// </summary>
        public int FavorCoherencyFailureOverNotFound { get; set; }

        /// <summary>
        /// Gets or sets A - Abort Remaining Put Changes on Failure (1 bit): A bit that specifies if set to abort remaining put changes on failure.
        /// </summary>
        public int AbortRemainingPutChangesOnFailure { get; set; }

        /// <summary>
        /// Gets or sets H - Multi-Request Put Hint (1 bit): A bit that specifies to reduce the number of auto coalesces during multi-request put scenarios, if only one request for a put changes, this bit is 0. 
        /// </summary>
        public int Reserved1Bit { get; set; }

        /// <summary>
        /// Gets or sets C - Return Complete Knowledge If Possible (1 bit): A bit that specifies to return the complete knowledge from the server provided that this request has exclusive access to the knowledge. Exclusive knowledge access is only granted on Coalesce and therefore complete knowledge will not be returned in non-coalescing sub-requests. 
        /// </summary>
        public int ReturnCompleteKnowledgeIfPossible { get; set; }

        /// <summary>
        /// Gets or sets LastWriterWinsOnNextChange (1 bit): A bit that specifies to allow the Put Changes to be subsequently overwritten on the next put changes.
        /// </summary>
        public int LastWriterWinsOnNextChange { get; set; }

        /// <summary>
        /// Gets or sets ContentVersionCoherencyCheck (variable): A Binary Item (section 2.2.1.3) which MUST be ignored..
        /// </summary>
        public BinaryItem ContentVersionCoherencyCheck { get; set; }

        /// <summary>
        /// Gets or sets Author Logins (variable): A String Item Array (section 2.2.1.14) structure that defines author login information.
        /// </summary>
        public StringItemArray AuthorLogins { get; set; }

        /// <summary>
        /// Gets or sets Reserved (1 byte): MUST be ignored
        /// </summary>
        public int Reserve1Byte { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the additional flags is used.
        /// </summary>
        public bool IsAdditionalFlagsUsed { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the diagnostic request option input is used.
        /// </summary>
        public bool IsDiagnosticRequestOptionInputUsed { get; set; }

        /// <summary>
        /// Gets or sets a value: Force Revision Chain Optimization (1 bit): A bit that specifies that the server should optimize the chain of revisions by refactoring them as part of the Put Changes request.
        /// If the IsDiagnosticRequestOptionInputUsed is false, this property will be ignored.
        /// </summary>
        public int ForceRevisionChainOptimization { get; set; }

        /// <summary>
        /// Gets or sets a value: Return Applied Storage Index Id Entries (1 bit): A bit that specifies that the storage indexes that are applied to the storage as part of the Put Changes will be returned in a Storage Index specified in the Put Changes Response by the Applied Storage Index Id.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int ReturnAppliedStorageIndexIdEntries { get; set; }

        /// <summary>
        /// Gets or sets a value: Return Data Elements Added (1 bit): A bit that specifies that the Data Elements that are added to the storage as part of this Put Changes will be return in a Data Element Collection specified in the Put Changes Response by the Data Elements Added collection.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int ReturnDataElementsAdded { get; set; }

        /// <summary>
        /// Gets or sets a value: Check for Id Reuse (1 bit): A bit that specifies that the server should attempt to check the Put Changes Request for the re-use of previously used Ids. This may occur when ID allocations are used and a client rollback occurs.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int CheckForIdReuse { get; set; }

        /// <summary>
        /// Gets or sets a value: Coherency Check Only Applied Index Entries (1 bit): A bit that specifies that only the index entries that are actually applied as part of the change will be checked for coherency.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int CoherencyCheckOnlyAppliedIndexEntries { get; set; }

        /// <summary>
        /// Gets or sets a value: A bit that specifies that the Put Changes request should be treated as a full file save with no dependencies on any pre-existing state. 
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int FullFileReplacePut { get; set; }

        /// <summary>
        /// Gets or sets a value: A bit that specifies that the Put Changes request will fail coherency if any of the supplied Storage Indexes are unrooted. 
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int RequireStorageMappingsRooted { get; set; }

        /// <summary>
        /// Gets or sets a value:  An 10-bit reserved field that MUST be set to zero.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public int Reserve { get; set; }

        /// <summary>
        /// Gets or sets a value: A compact unsigned 64-bit integer (section 2.2.1.1) that MUST be ignored.
        /// If the IsAdditionalFlagsUsed is false, this property will be ignored.
        /// </summary>
        public Compact64bitInt Reserve2 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the lock id is used.
        /// </summary>
        public bool IsLockIdUsed { get; set; }

        /// <summary>
        /// Gets or sets the lock ID.
        /// If the IsLockIdUsed is false, this property will be ignored.
        /// </summary>
        public Guid LockID { get; set; }

        /// <summary>
        /// Gets or sets the optional client knowledge.
        /// </summary>
        public Knowledge OptionalClientKnowledge { get; set; }
        
        /// <summary>
        /// Gets or sets Put Changes Request (4 bytes): A stream object header that specifies a put changes request.
        /// </summary>
        internal StreamObjectHeaderStart32bit PutChangesRequestStart { get; set; }

        /// <summary>
        /// Gets or sets Additional Flags Header (4 bytes): A 32-bit stream object header that specifies the start of an Additional Flags structure.
        /// </summary>
        internal StreamObjectHeaderStart32bit AdditionalFlagsStart { get; set; }

        /// <summary>
        /// Gets or sets Diagnostic Request Option Input Header (4 bytes): A 32-bit stream object header that specifies the start of an Diagnostic Request Option Input structure.
        /// </summary>
        internal StreamObjectHeaderStart32bit DiagnosticRequestOptionInputStart { get; set; }

        /// <summary>
        /// Gets or sets Lock Id Header (4 bytes): A 32-bit stream object header that specifies a Lock Id start.
        /// </summary>
        internal StreamObjectHeaderStart32bit LockIdStart { get; set; }

        /// <summary>
        /// This method is used to convert the element into a byte List 
        /// </summary>
        /// <returns>Return the Byte List</returns>
        public override List<byte> SerializeToByteList()
        {
            // Storage Index Extended GUID
            this.StorageIndexExtendedGUID = this.StorageIndexExtendedGUID ?? new ExGuid();
            List<byte> storageIndexExtendedGUIDBytes = this.StorageIndexExtendedGUID.SerializeToByteList();

            // Expect Storage Index Extended GUID
            List<byte> expectedStorageIndexExtendedGUIDBytes = this.ExpectedStorageIndexExtendedGUID.SerializeToByteList();

            // ContentVersionCoherencyCheck
            List<byte> contentVersionCoherencyCheckBytes = this.ContentVersionCoherencyCheck.SerializeToByteList();

            // Author Logins 
            List<byte> authorLoginsBytes = this.AuthorLogins.SerializeToByteList();

            // Put Changes Request
            this.PutChangesRequestStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.PutChangesRequest, 1 + storageIndexExtendedGUIDBytes.Count + expectedStorageIndexExtendedGUIDBytes.Count + contentVersionCoherencyCheckBytes.Count + authorLoginsBytes.Count + 1);
            List<byte> putChangesRequestBytes = this.PutChangesRequestStart.SerializeToByteList();
            
            // reserved
            BitWriter bitWriter = new BitWriter(1);
            bitWriter.AppendInit32(this.ImplyNullExpectedIfNoMapping, 1);
            bitWriter.AppendInit32(this.Partial, 1);
            bitWriter.AppendInit32(this.PartialLast, 1);
            bitWriter.AppendInit32(this.FavorCoherencyFailureOverNotFound, 1);
            bitWriter.AppendInit32(this.AbortRemainingPutChangesOnFailure, 1);
            bitWriter.AppendInit32(this.Reserved1Bit, 1);
            bitWriter.AppendInit32(this.ReturnCompleteKnowledgeIfPossible, 1);
            bitWriter.AppendInit32(this.LastWriterWinsOnNextChange, 1);          

            // Reserve1Byte
            List<byte> reserve1ByteBytes = new List<byte>(new byte[1]);

            List<byte> reservedBytes = new List<byte>(bitWriter.Bytes);

            List<byte> byteList = new List<byte>();

            // sub-request start
            byteList.AddRange(base.SerializeToByteList());
            
            // put change request
            byteList.AddRange(putChangesRequestBytes);
            
            // Storage Index Extended GUID
            byteList.AddRange(storageIndexExtendedGUIDBytes);

            // Expected Storage Index Extended GUID
            byteList.AddRange(expectedStorageIndexExtendedGUIDBytes);
            
            // reserved
            byteList.AddRange(reservedBytes);

            // ContentVersionCoherencyCheck
            byteList.AddRange(contentVersionCoherencyCheckBytes);

            // Author Logins
            byteList.AddRange(authorLoginsBytes);

            // Reserve1Byte
            byteList.AddRange(reserve1ByteBytes);

            if (this.IsAdditionalFlagsUsed)
            {
                this.AdditionalFlagsStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.AdditionalFlags, 2);
                byteList.AddRange(this.AdditionalFlagsStart.SerializeToByteList());

                BitWriter additionalFlagsWriter = new BitWriter(2);
                additionalFlagsWriter.AppendInit32(this.ReturnAppliedStorageIndexIdEntries, 1);
                additionalFlagsWriter.AppendInit32(this.ReturnDataElementsAdded, 1);
                additionalFlagsWriter.AppendInit32(this.CheckForIdReuse, 1);
                additionalFlagsWriter.AppendInit32(this.CoherencyCheckOnlyAppliedIndexEntries, 1);
                additionalFlagsWriter.AppendInit32(this.FullFileReplacePut, 1);
                additionalFlagsWriter.AppendInit32(this.RequireStorageMappingsRooted, 1);
                additionalFlagsWriter.AppendInit32(this.Reserve, 10);              
                byteList.AddRange(additionalFlagsWriter.Bytes);

                this.Reserve2 = new Compact64bitInt(0x0002000000000000);
                byteList.AddRange(this.Reserve2.SerializeToByteList());

            }

            if (this.IsLockIdUsed)
            {
                this.LockIdStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.PutChangesLockId, 16);
                byteList.AddRange(this.LockIdStart.SerializeToByteList());
                byteList.AddRange(this.LockID.ToByteArray());
            }

            if (this.OptionalClientKnowledge != null)
            {
                byteList.AddRange(this.OptionalClientKnowledge.SerializeToByteList());
            }

            if (this.IsDiagnosticRequestOptionInputUsed)
            {
                this.DiagnosticRequestOptionInputStart = new StreamObjectHeaderStart32bit(StreamObjectTypeHeaderStart.DiagnosticRequestOptionInput, 2);
                byteList.AddRange(this.DiagnosticRequestOptionInputStart.SerializeToByteList());

                BitWriter diagnosticRequestOptionWriter = new BitWriter(2);
                diagnosticRequestOptionWriter.AppendInit32(this.ForceRevisionChainOptimization, 1);
                diagnosticRequestOptionWriter.AppendInit32(this.Reserve, 7);
                byteList.AddRange(diagnosticRequestOptionWriter.Bytes);
            }

            // sub-request end
            byteList.AddRange(this.ToBytesEnd());

            return byteList;
        }
    }
}
