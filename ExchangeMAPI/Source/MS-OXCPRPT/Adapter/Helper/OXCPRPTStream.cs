namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Implementation for MS_OXCPRPTAdapter for stream operations
    /// </summary>
    public partial class MS_OXCPRPTAdapter : ManagedAdapterBase, IMS_OXCPRPTAdapter
    {
        /// <summary>
        /// This ROP opens a property for streaming access.
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="openStreamResponse">The response of this ROP.</param>
        /// <param name="tag">The propertyTag structure.</param>
        /// <param name="openMode">8-bit flags structure. These flags control how the stream is opened.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The handle returned.</returns>
        private uint RopOpenStream(uint objHandle, out RopOpenStreamResponse openStreamResponse, PropertyTag tag, byte openMode, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopOpenStreamRequest openStreamRequest;

            openStreamRequest.RopId = (byte)RopId.RopOpenStream;
            openStreamRequest.LogonId = LogonId;
            openStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            openStreamRequest.OutputHandleIndex = (byte)HandleIndex.SecondIndex;
            openStreamRequest.PropertyTag = tag;
            openStreamRequest.OpenModeFlags = openMode;

            this.responseSOHsValue = this.ProcessSingleRop(openStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            openStreamResponse = (RopOpenStreamResponse)this.responseValue;
            uint streamObjectHandle = this.responseSOHsValue[0][openStreamResponse.OutputHandleIndex];
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, openStreamResponse.ReturnValue, string.Format("RopOpenStream Failed! Error: 0x{0:X8}", openStreamResponse.ReturnValue));
            }

            return streamObjectHandle;
        }

        /// <summary>
        /// This ROP reads bytes from a stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="byteCount">The maximum number of bytes to read if the value is not equal to 0xBABE.</param>
        /// <param name="maximumByteCount">The maximum number of bytes to read if the value of the ByteCount field is equal to 0xBABE.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopReadStreamResponse RopReadStream(uint objHandle, ushort byteCount, uint maximumByteCount, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopReadStreamRequest readStreamRequest;
            RopReadStreamResponse readStreamResponse;

            readStreamRequest.RopId = (byte)RopId.RopReadStream;
            readStreamRequest.LogonId = LogonId;
            readStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            readStreamRequest.ByteCount = byteCount;
            readStreamRequest.MaximumByteCount = maximumByteCount;

            this.responseSOHsValue = this.ProcessSingleRop(readStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            readStreamResponse = (RopReadStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, readStreamResponse.ReturnValue, string.Format("RopReadStream failed! Error: 0x{0:X8}", readStreamResponse.ReturnValue));
            }

            return readStreamResponse;
        }

        /// <summary>
        /// This ROP writes bytes to a stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="writeData">These values specify the data how to be written to the stream.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopWriteStreamResponse RopWriteStream(uint objHandle, string writeData, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopWriteStreamRequest writeStreamRequest;
            RopWriteStreamResponse writeStreamResponse;

            writeStreamRequest.RopId = (byte)RopId.RopWriteStream;
            writeStreamRequest.LogonId = LogonId;
            writeStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;

            byte[] data = Encoding.ASCII.GetBytes(writeData);
            writeStreamRequest.DataSize = (ushort)data.Length;
            writeStreamRequest.Data = data;

            this.responseSOHsValue = this.ProcessSingleRop(writeStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            writeStreamResponse = (RopWriteStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, writeStreamResponse.ReturnValue, string.Format("RopWriteStream failed! Error: 0x{0:X8}", writeStreamResponse.ReturnValue));
            }

            return writeStreamResponse;
        }

        /// <summary>
        /// This ROP commits stream operations. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopCommitStreamResponse RopCommitStream(uint objHandle, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCommitStreamRequest commitStreamRequest;
            RopCommitStreamResponse commitStreamResponse;

            commitStreamRequest.RopId = (byte)RopId.RopCommitStream;
            commitStreamRequest.LogonId = LogonId;
            commitStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;

            this.responseSOHsValue = this.ProcessSingleRop(commitStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            commitStreamResponse = (RopCommitStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, commitStreamResponse.ReturnValue, string.Format("RopCommitStream failed! Error: 0x{0:X8}", commitStreamResponse.ReturnValue));
            }

            return commitStreamResponse;
        }

        /// <summary>
        /// This ROP gets the size of a stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopGetStreamSizeResponse RopGetStreamSize(uint objHandle, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetStreamSizeRequest getStreamSizeRequest;
            RopGetStreamSizeResponse getStreamSizeResponse;

            getStreamSizeRequest.RopId = (byte)RopId.RopGetStreamSize;
            getStreamSizeRequest.LogonId = LogonId;
            getStreamSizeRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;

            this.responseSOHsValue = this.ProcessSingleRop(getStreamSizeRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getStreamSizeResponse = (RopGetStreamSizeResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, getStreamSizeResponse.ReturnValue, string.Format("RopGetStreamSize failed! Error: 0x{0:X8}", getStreamSizeResponse.ReturnValue));
            }

            return getStreamSizeResponse;
        }

        /// <summary>
        /// This ROP sets the size of a stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="streamSize">The size of the stream.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopSetStreamSizeResponse RopSetStreamSize(uint objHandle, ulong streamSize, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSetStreamSizeRequest setStreamSizeRequest;
            RopSetStreamSizeResponse setStreamSizeResponse;

            setStreamSizeRequest.RopId = (byte)RopId.RopSetStreamSize;
            setStreamSizeRequest.LogonId = LogonId;
            setStreamSizeRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            setStreamSizeRequest.StreamSize = streamSize;

            this.responseSOHsValue = this.ProcessSingleRop(setStreamSizeRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            setStreamSizeResponse = (RopSetStreamSizeResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, setStreamSizeResponse.ReturnValue, string.Format("RopSetStreamSize failed! Error: 0x{0:X8}", setStreamSizeResponse.ReturnValue));
            }

            return setStreamSizeResponse;
        }

        /// <summary>
        /// This ROP seeks to a specific offset within a stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="origin">The origin location for the seek operation.</param>
        /// <param name="offset">The seek offset for the seek operation.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopSeekStreamResponse RopSeekStream(uint objHandle, byte origin, long offset, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSeekStreamRequest seekStreamRequest;
            RopSeekStreamResponse seekStreamResponse;

            seekStreamRequest.RopId = (byte)RopId.RopSeekStream;
            seekStreamRequest.LogonId = LogonId;
            seekStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            seekStreamRequest.Origin = origin;
            seekStreamRequest.Offset = (ulong)offset;

            this.responseSOHsValue = this.ProcessSingleRop(seekStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            seekStreamResponse = (RopSeekStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, seekStreamResponse.ReturnValue, string.Format("RopSeekStream failed! Error: 0x{0:X8}", seekStreamResponse.ReturnValue));
            }

            return seekStreamResponse;
        }

        /// <summary>
        /// This ROP copies a number of bytes from a source stream to a destination stream.
        /// </summary>
        /// <param name="sourceHandle">The source handle to copy.</param>
        /// <param name="destHandle">The destination handle to be copied.</param>
        /// <param name="sourceHandleIndex">The source object stored location index in Server object handle table.</param>
        /// <param name="destHandleIndex">This destination object stored location index in the Server object handle table.</param>
        /// <param name="byteCount">The number of bytes to be copied.</param>
        /// <returns>The response of this ROP.</returns>
        private RopCopyToStreamResponse RopCopyToStream(uint sourceHandle, uint destHandle, byte sourceHandleIndex, byte destHandleIndex, ulong byteCount)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCopyToStreamRequest copyToStreamRequest;
            RopCopyToStreamResponse copyToStreamResponse;

            copyToStreamRequest.RopId = (byte)RopId.RopCopyToStream;
            copyToStreamRequest.LogonId = LogonId;
            copyToStreamRequest.SourceHandleIndex = sourceHandleIndex;
            copyToStreamRequest.DestHandleIndex = destHandleIndex;
            copyToStreamRequest.ByteCount = byteCount;

            List<uint> handleList = new List<uint>
            {
                sourceHandle, destHandle
            };

            this.responseSOHsValue = this.ProcessSingleRopWithMutipleServerObjects(copyToStreamRequest, handleList, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            copyToStreamResponse = (RopCopyToStreamResponse)this.responseValue;
            return copyToStreamResponse;
        }

        /// <summary>
        /// This ROP locks a specified range of bytes in a stream.  
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="regionOffset">The byte location in the stream where the region begins.</param>
        /// <param name="regionSize">The size of the region, in bytes.</param>
        /// <param name="lockFlags">The flags specified the behavior of the lock operation.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopLockRegionStreamResponse RopLockRegionStream(uint objHandle, ulong regionOffset, ulong regionSize, uint lockFlags, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopLockRegionStreamRequest lockRegionStreamRequest;
            RopLockRegionStreamResponse lockRegionStreamResponse;

            lockRegionStreamRequest.RopId = (byte)RopId.RopLockRegionStream;
            lockRegionStreamRequest.LogonId = LogonId;
            lockRegionStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            lockRegionStreamRequest.RegionOffset = regionOffset;
            lockRegionStreamRequest.RegionSize = regionSize;
            lockRegionStreamRequest.LockFlags = lockFlags;
            this.responseSOHsValue = this.ProcessSingleRop(lockRegionStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            lockRegionStreamResponse = (RopLockRegionStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, lockRegionStreamResponse.ReturnValue, string.Format("RopLockRegionStream failed! Error: 0x{0:X8}", lockRegionStreamResponse.ReturnValue));
            }

            return lockRegionStreamResponse;
        }

        /// <summary>
        /// This ROP unlocks a specified range of bytes in a stream.  
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="regionOffset">The byte location in the stream where the region begins.</param>
        /// <param name="regionSize">The size of the region in bytes.</param>
        /// <param name="lockFlags">The flag specified the behavior of the lock operation.</param>
        /// <returns>The response of this ROP.</returns>
        private RopUnlockRegionStreamResponse RopUnlockRegionStream(uint objHandle, ulong regionOffset, ulong regionSize, uint lockFlags)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopUnlockRegionStreamRequest unlockRegionStreamRequest;
            RopUnlockRegionStreamResponse unlockRegionStreamResponse;

            unlockRegionStreamRequest.RopId = (byte)RopId.RopUnlockRegionStream;
            unlockRegionStreamRequest.LogonId = LogonId;
            unlockRegionStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            unlockRegionStreamRequest.RegionOffset = regionOffset;
            unlockRegionStreamRequest.RegionSize = regionSize;
            unlockRegionStreamRequest.LockFlags = lockFlags;

            this.responseSOHsValue = this.ProcessSingleRop(unlockRegionStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            unlockRegionStreamResponse = (RopUnlockRegionStreamResponse)this.responseValue;
            return unlockRegionStreamResponse;
        }

        /// <summary>
        /// This ROP writes bytes to a stream and commits the stream. 
        /// </summary>
        /// <param name="objHandle">The handle to operate.</param>
        /// <param name="writeData">The data to be written to the stream.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopWriteStreamResponse RopWriteAndCommitStream(uint objHandle, string writeData, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopWriteAndCommitStreamRequest writeAndCommitStreamRequest;
            RopWriteStreamResponse writeAndCommitStreamResponse;

            writeAndCommitStreamRequest.RopId = (byte)RopId.RopWriteAndCommitStream;
            writeAndCommitStreamRequest.LogonId = LogonId;
            writeAndCommitStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            byte[] data = Encoding.ASCII.GetBytes(writeData);
            writeAndCommitStreamRequest.DataSize = (ushort)data.Length;
            writeAndCommitStreamRequest.Data = data;

            this.responseSOHsValue = this.ProcessSingleRop(writeAndCommitStreamRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            writeAndCommitStreamResponse = (RopWriteStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, writeAndCommitStreamResponse.ReturnValue, string.Format("RopWriteAndCommitStream failed! Error: 0x{0:X8}", writeAndCommitStreamResponse.ReturnValue));
            }

            return writeAndCommitStreamResponse;
        }

        /// <summary>
        /// This ROP creates a new stream object based on the same data as another stream. 
        /// </summary>
        /// <param name="sourceHandle">The source handle.</param>
        /// <param name="destHandle">The destination handle.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopCloneStreamResponse RopCloneStream(uint sourceHandle, uint destHandle, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCloneStreamRequest cloneStreamRequest;
            RopCloneStreamResponse cloneStreamResponse;

            cloneStreamRequest.RopId = (byte)RopId.RopCloneStream;
            cloneStreamRequest.LogonId = LogonId;
            cloneStreamRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            cloneStreamRequest.OutputHandleIndex = (byte)HandleIndex.SecondIndex;

            List<uint> handleList = new List<uint>
            {
                sourceHandle, destHandle
            };

            this.responseSOHsValue = this.ProcessSingleRopWithMutipleServerObjects(cloneStreamRequest, handleList, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            cloneStreamResponse = (RopCloneStreamResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, cloneStreamResponse.ReturnValue, string.Format("RopCloneStream failed! Error: 0x{0:X8}", cloneStreamResponse.ReturnValue));
            }

            return cloneStreamResponse;
        }
    }
}