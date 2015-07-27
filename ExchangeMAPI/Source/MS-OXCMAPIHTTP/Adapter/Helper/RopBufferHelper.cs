//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to build the request payload and parse the response payload.
    /// </summary>
    public class RopBufferHelper
    {
        /// <summary>
        /// This is used to set Flags of RPC_HEADER_EXT,which indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT.
        /// </summary>
        private const ushort LastForFlagsOfHeader = 0x0004;

        /// <summary>
        /// Length of RPC_HEADER_EXT
        /// </summary>
        private static readonly int RPCHEADEREXTLEN = Marshal.SizeOf(typeof(RPC_HEADER_EXT));

        /// <summary>
        /// This is the test site.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// Initializes a new instance of the RopBufferHelper class.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite.</param>
        public RopBufferHelper(ITestSite testSite)
        {
            this.site = testSite;
        }

        /// <summary>
        /// The method creates a request without ROPs.
        /// </summary>
        /// <returns>The ROPs request.</returns>
        public byte[] BuildRequestBufferWithoutRop()
        {
            // Only RopSize field (2 bytes) exists in the payload of request without ROPs.
            int payloadLen = 2;
            ushort ropSize = 2;
            int rpcHeaderExtlength = Marshal.SizeOf(typeof(RPC_HEADER_EXT));

            byte[] requestBuffer = new byte[rpcHeaderExtlength + payloadLen];
            int index = 0;

            // Constructs RPC header ext buffer
            RPC_HEADER_EXT rpcHeaderExt = new RPC_HEADER_EXT
            {
                // According to Open Specification, Version filed MUST be set to 0x0000.
                Version = 0,
                Flags = (byte)RpcHeaderExtFlags.Last,
                Size = (ushort)payloadLen,
                SizeActual = (ushort)payloadLen
            };

            // No compression and No obfuscation
            IntPtr ptr = Marshal.AllocHGlobal(rpcHeaderExtlength);

            // Release ptr in final sub-statement to make sure the resources will be released even if an exception occurs
            try
            {
                Marshal.StructureToPtr(rpcHeaderExt, ptr, true);
                Marshal.Copy(ptr, requestBuffer, index, rpcHeaderExtlength);
                index += rpcHeaderExtlength;
            }
            catch (ArgumentException e)
            {
                // Marshal.StructureToPtr or Marshal.Copy throw exception.
                throw new NotImplementedException(e.Message);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }

            Array.Copy(BitConverter.GetBytes(ropSize), 0, requestBuffer, index, sizeof(ushort));

            return requestBuffer;
        }

        /// <summary>
        /// The method creates single ROP request buffer.
        /// </summary>
        /// <param name="requestROP">ROP request object.</param>
        /// <param name="insideObjectHandle">Server object handle in request.</param>
        /// <returns>The ROP request buffer.</returns>
        public byte[] BuildRequestBuffer(ISerializable requestROP, uint insideObjectHandle)
        {
            List<ISerializable> requestRops = new List<ISerializable>
            {
                requestROP
            };

            List<uint> requestSOH = new List<uint>
            {
                insideObjectHandle
            };

            if (Common.IsOutputHandleInRopRequest(requestROP))
            {
                // Add an element for server output object handle, set default value to 0xFFFFFFFF
                requestSOH.Add(ConstValues.DefaultOutputHandle);
            }

            return this.BuildRequestBuffer(requestRops, requestSOH);
        }

        /// <summary>
        /// The method creates ROPs request buffer.
        /// </summary>
        /// <param name="requestROPs">ROPs in request.</param>
        /// <param name="requestSOHTable">Server object handles table.</param>
        /// <returns>The ROPs request buffer.</returns>
        public byte[] BuildRequestBuffer(List<ISerializable> requestROPs, List<uint> requestSOHTable)
        {
            // Definition for PayloadLen which indicates the length of the field that represents the length of payload.
            int payloadLen = 0x2;
            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    payloadLen += requestROP.Size();
                }
            }

            ushort ropSize = (ushort)payloadLen;

            if (requestSOHTable != null)
            {
                payloadLen += requestSOHTable.Count * sizeof(uint);
            }

            byte[] requestBuffer = new byte[RPCHEADEREXTLEN + payloadLen];
            int index = 0;

            // Construct RPC header ext buffer
            RPC_HEADER_EXT rpcHeaderExt = new RPC_HEADER_EXT();

            // There is only one version of the header at this time so this value MUST be set to 0x00.
            rpcHeaderExt.Version = 0x00;

            // Last (0x04) indicates that no other RPC_HEADER_EXT follows the data of the current RPC_HEADER_EXT. 
            rpcHeaderExt.Flags = LastForFlagsOfHeader;

            rpcHeaderExt.Size = (ushort)payloadLen;
            rpcHeaderExt.SizeActual = rpcHeaderExt.Size;

            IntPtr ptr = Marshal.AllocHGlobal(RPCHEADEREXTLEN);
            try
            {
                Marshal.StructureToPtr(rpcHeaderExt, ptr, true);
                Marshal.Copy(ptr, requestBuffer, index, RPCHEADEREXTLEN);
                index += RPCHEADEREXTLEN;
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }

            // RopSize's type is ushort. So the offset will be 2.
            Array.Copy(BitConverter.GetBytes(ropSize), 0, requestBuffer, index, 2);
            index += 2;

            if (requestROPs != null)
            {
                foreach (ISerializable requestROP in requestROPs)
                {
                    Array.Copy(requestROP.Serialize(), 0, requestBuffer, index, requestROP.Size());
                    index += requestROP.Size();
                }
            }

            if (requestSOHTable != null)
            {
                foreach (uint serverHandle in requestSOHTable)
                {
                    Array.Copy(BitConverter.GetBytes(serverHandle), 0, requestBuffer, index, sizeof(uint));
                    index += sizeof(uint);
                }
            }

            // Compress and obfuscate request buffer as configured.
            requestBuffer = Common.CompressAndObfuscateRequest(requestBuffer, this.site);

            return requestBuffer;
        }

        /// <summary>
        /// The method parses response buffer.
        /// </summary>
        /// <param name="rgbOut">The ROP response payload.</param>
        /// <param name="rpcHeaderExts">RPC header ext.</param>
        /// <param name="rops">ROPs in response.</param>
        /// <param name="serverHandleObjectsTables">Server handle objects tables</param>
        public void ParseResponseBuffer(byte[] rgbOut, out RPC_HEADER_EXT[] rpcHeaderExts, out byte[][] rops, out uint[][] serverHandleObjectsTables)
        {
            List<RPC_HEADER_EXT> rHE = new List<RPC_HEADER_EXT>();
            List<byte[]> ropList = new List<byte[]>();
            List<uint[]> shot = new List<uint[]>();
            IntPtr ptr = IntPtr.Zero;

            int index = 0;
            bool end = false;
            do
            {
                // Parse rpc header ext
                RPC_HEADER_EXT rpcHeaderExt;
                ptr = Marshal.AllocHGlobal(RPCHEADEREXTLEN);
                try
                {
                    Marshal.Copy(rgbOut, index, ptr, RPCHEADEREXTLEN);
                    rpcHeaderExt = (RPC_HEADER_EXT)Marshal.PtrToStructure(ptr, typeof(RPC_HEADER_EXT));
                    end = (rpcHeaderExt.Flags & LastForFlagsOfHeader) == LastForFlagsOfHeader;
                    rHE.Add(rpcHeaderExt);
                    index += RPCHEADEREXTLEN;
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }

                // Parse ropSize
                ushort ropSize = BitConverter.ToUInt16(rgbOut, index);
                index += sizeof(ushort);

                if ((ropSize - sizeof(ushort)) > 0)
                {
                    // Parse ROP
                    byte[] rop = new byte[ropSize - sizeof(ushort)];
                    Array.Copy(rgbOut, index, rop, 0, ropSize - sizeof(ushort));
                    ropList.Add(rop);
                    index += ropSize - sizeof(ushort);
                }

                // Parse server handle objects table
                this.site.Assert.IsTrue((rpcHeaderExt.Size - ropSize) % sizeof(uint) == 0, "server object handle should be uint32 array");

                int count = (rpcHeaderExt.Size - ropSize) / sizeof(uint);
                if (count > 0)
                {
                    uint[] sohs = new uint[count];
                    for (int k = 0; k < count; k++)
                    {
                        sohs[k] = BitConverter.ToUInt32(rgbOut, index);
                        index += sizeof(uint);
                    }

                    shot.Add(sohs);
                }
            }
            while (!end);

            rpcHeaderExts = rHE.ToArray();
            rops = ropList.ToArray();
            serverHandleObjectsTables = shot.ToArray();
        }
    }
}