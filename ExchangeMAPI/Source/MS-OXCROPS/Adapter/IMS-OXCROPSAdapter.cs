namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// IMS_OXCROPSAdapter is the protocol adapter interface for MS-OXCROPS.
    /// </summary>
    public interface IMS_OXCROPSAdapter : IAdapter
    {
        /// <summary>
        /// Connect to the server for RPC calling.
        /// </summary>
        /// <param name="server">Server to connect.</param>
        /// <param name="connectionType">the type of connection</param>
        /// <param name="userDN">User DN used to connect server</param>
        /// <param name="domain">The domain the server is deployed</param>
        /// <param name="userName">The domain account name</param>
        /// <param name="password">User Password</param>
        /// <returns>Result of connecting.</returns>
        bool RpcConnect(string server, ConnectionType connectionType, string userDN, string domain, string userName, string password);

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>Result of disconnecting.</returns>
        bool RpcDisconnect();

        /// <summary>
        /// set auto redirect value in RPC context
        /// If setting this to true, the RPC server will return EcWrongServer error (0x478). And the request will be redirected to designated server.
        /// If setting this to false, the RPC server will return EcWrongServer error (0x478). But the request will not be redirected.
        /// </summary>
        /// <param name="option">true indicates enable auto redirect, false indicates disable auto redirect</param>
        void SetAutoRedirect(bool option);

        /// <summary>
        /// Method which executes single ROP.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> ProcessSingleRop(
            ISerializable ropRequest, 
            uint inputObjHandle, 
            ref IDeserializable response, 
            ref byte[] rawData, 
            RopResponseType expectedRopResponseType);

        /// <summary>
        /// Method which executes single ROP.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <param name="returnValue">The return value of the ROP method.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> ProcessSingleRopWithReturnValue(
            ISerializable ropRequest,
            uint inputObjHandle,
            ref IDeserializable response,
            ref byte[] rawData,
            RopResponseType expectedRopResponseType,
            out uint returnValue);

        /// <summary>
        /// Method which executes single ROP operation with the maximum size of the rgbOut buffer set as pcbOut.
        /// For more detail about rgbOut and pcbOut, see [MS-OXCRPC].
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="inputObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <param name="pcbOut">The maximum size of the rgbOut buffer to place Response in.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> ProcessSingleRopWithOptionResponseBufferSize(
            ISerializable ropRequest, 
            uint inputObjHandle, 
            ref IDeserializable response, 
            ref byte[] rawData, 
            RopResponseType expectedRopResponseType, 
            uint pcbOut);

        /// <summary>
        /// Method which executes multiple ROPs.
        /// </summary>
        /// <param name="requestRops">ROP request objects.</param>
        /// <param name="inputObjHandles">Server object handles in request.</param>
        /// <param name="responseRops">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">The expected response type.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> ProcessMutipleRops(
            List<ISerializable> requestRops, 
            List<uint> inputObjHandles, 
            ref List<IDeserializable> responseRops, 
            ref byte[] rawData,
            RopResponseType expectedRopResponseType);

        /// <summary>
        /// Method which executes single ROP with multiple server objects.
        /// </summary>
        /// <param name="ropRequest">ROP request object.</param>
        /// <param name="inputObjHandles">Server object handles in request.</param>
        /// <param name="response">ROP response object.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="expectedRopResponseType">ROP response type expected.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> ProcessSingleRopWithMutipleServerObjects(
            ISerializable ropRequest, 
            List<uint> inputObjHandles, 
            ref IDeserializable response, 
            ref byte[] rawData, 
            RopResponseType expectedRopResponseType);
    }
}