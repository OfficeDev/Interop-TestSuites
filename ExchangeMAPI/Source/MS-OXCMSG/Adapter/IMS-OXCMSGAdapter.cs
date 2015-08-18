namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The MS-OXCMSG protocol adapter interface.
    /// </summary>
    public interface IMS_OXCMSGAdapter : IAdapter
    {
        /// <summary>
        /// Connect to the server.
        /// </summary>
        /// <param name="connectionType">The type of connection</param>
        /// <param name="user">A string value indicates the domain account name that connects to server.</param>
        /// <param name="password">A string value indicates the password of the user which is used.</param>
        /// <param name="userDN">A string that identifies user who is making the EcDoConnectEx call</param>
        /// <returns>A Boolean value indicating whether connects successfully.</returns>
        bool RpcConnect(ConnectionType connectionType, string user, string password, string userDN);

        /// <summary>
        /// Disconnect from the server.
        /// </summary>
        /// <returns>A Boolean value indicating whether disconnects successfully.</returns>
        bool RpcDisconnect();

        /// <summary>
        /// Send ROP request with single operation.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="getPropertiesFlags">The flag indicate the test cases expect to get which object type's properties(message's properties or attachment's properties).</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object response, ref byte[] rawData, GetPropertiesFlags getPropertiesFlags);

        /// <summary>
        /// Send ROP request with single operation.
        /// </summary>
        /// <param name="ropRequest">ROP request objects.</param>
        /// <param name="insideObjHandle">Server object handle in request.</param>
        /// <param name="response">ROP response objects.</param>
        /// <param name="rawData">The ROP response payload.</param>
        /// <param name="getPropertiesFlags">The flag indicate the test cases expect to get which object type's properties(message's properties or attachment's properties).</param>
        /// <param name="returnValue">An unsigned integer value indicates the return value of call EcDoRpcExt2 method.</param>
        /// <returns>Server objects handles in response.</returns>
        List<List<uint>> DoRopCall(ISerializable ropRequest, uint insideObjHandle, ref object response, ref byte[] rawData, GetPropertiesFlags getPropertiesFlags, out uint returnValue);

        /// <summary>
        /// Get the named properties value of specified Message object.
        /// </summary>
        /// <param name="longIdProperties">The list of named properties</param>
        /// <param name="messageHandle">The object handle of specified Message object.</param>
        /// <returns>Returns named property values of specified Message object.</returns>
        Dictionary<PropertyNames, byte[]> GetNamedPropertyValues(List<PropertyNameObject> longIdProperties, uint messageHandle);
    }
}