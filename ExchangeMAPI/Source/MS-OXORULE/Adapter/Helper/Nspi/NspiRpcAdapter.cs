namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The NspiRpcAdapter class contains the RPC implements for the interfaces of IMS_OXNSPIAdapter.
    /// </summary>
    public class NspiRpcAdapter : ManagedAdapterBase
    {
        #region Variable

        /// <summary>
        /// The Site instance.
        /// </summary>
        private ITestSite site;

        /// <summary>
        /// The RPC binding.
        /// </summary>
        private IntPtr rpcBinding = IntPtr.Zero;

        /// <summary>
        /// The RPC context handle.
        /// </summary>
        private IntPtr contextHandle = IntPtr.Zero;

        /// <summary>
        /// The time internal that is used to wait to retry when the returned error code is GeneralFailure.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// The retry count that is used to retry when the returned error code is GeneralFailure.
        /// </summary>
        private uint maxRetryCount;
        #endregion
       
        /// <summary>
        /// Initializes a new instance of the NspiRpcAdapter class.
        /// </summary>
        /// <param name="site">The site instance.</param>
        /// <param name="rpcBinding">The RPC binding.</param>
        /// <param name="contextHandle">The RPC context handle.</param>
        /// <param name="waitTime">The time internal that is used to wait to retry when the returned error code is GeneralFailure.</param>
        /// <param name="maxRetryCount">The retry count that is used to retry when the returned error code is GeneralFailure.</param>
        public NspiRpcAdapter(ITestSite site, IntPtr rpcBinding, IntPtr contextHandle, int waitTime, uint maxRetryCount)
        {
            this.site = site;
            this.rpcBinding = rpcBinding;
            this.contextHandle = contextHandle;
            this.waitTime = waitTime;
            this.maxRetryCount = maxRetryCount;
        }

        #region Instance interface

        /// <summary>
        /// The NspiBind method initiates a session between a client and the server.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="serverGuid">The value NULL or a pointer to a GUID value that is associated with the specific server.</param>
        /// <param name="needRetry">A Boolean value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiBind(uint flags, STAT stat, ref FlatUID_r? serverGuid, bool needRetry = true)
        {
            int result;

            IntPtr ptrServerGuid = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);

            if (serverGuid.HasValue && serverGuid.Value.Ab != null)
            {
                ptrServerGuid = AdapterHelper.AllocFlatUID_r(serverGuid.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiBind(this.rpcBinding, flags, ref stat, ptrServerGuid, ref this.contextHandle);
                }
                catch (SEHException e)
                {
                    result = (int)NativeMethods.RpcExceptionCode(e);
                    this.site.Log.Add(LogEntryKind.Comment, "RPC component throws exception, the error code is {0}, the error message is: {1}", result, new Win32Exception(result).ToString());
                }

                if ((ErrorCodeValue)result == ErrorCodeValue.GeneralFailure && needRetry)
                {
                    Thread.Sleep(this.waitTime);
                }
                else
                {
                    break;
                }

                retryCount++;
            }
            while ((ErrorCodeValue)result == ErrorCodeValue.GeneralFailure && retryCount < this.maxRetryCount);

            if (!Enum.IsDefined(typeof(ErrorCodeValue), (uint)result))
            {
                throw new ArgumentException(string.Format("An unknown error is returned, the error code is: {0} and the error message is: {1}", result, new Win32Exception(result).ToString()));
            }

            // Parse ServerGuid from ptrServerGuid.
            if (ptrServerGuid == IntPtr.Zero)
            {
                serverGuid = null;
            }
            else
            {
                serverGuid = AdapterHelper.ParseFlatUID_r(ptrServerGuid);
            }

            // Free allocated memory for serverGuid.
            if (serverGuid.HasValue)
            {
                if (serverGuid.Value.Ab != null)
                {
                    Marshal.FreeHGlobal(ptrServerGuid);
                }
            }

            // Free allocated memory for stat.
            Marshal.FreeHGlobal(ptrStat);
            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiUnbind method destroys the context handle. No other action is taken.
        /// </summary>
        /// <param name="reserved">A DWORD [MS-DTYP] value reserved for future use. This property is ignored by the server.</param>
        /// <param name="contextHandle">The RPC context handle.</param>
        /// <returns>A DWORD value that specifies the return status of the method.</returns>
        public uint NspiUnbind(uint reserved, ref IntPtr contextHandle)
        {
            uint result;

            try
            {
                result = OxnspiInterop.NspiUnbind(ref this.contextHandle, reserved);
                contextHandle = this.contextHandle;
            }
            catch (SEHException e)
            {
                result = NativeMethods.RpcExceptionCode(e);
                this.site.Log.Add(LogEntryKind.Comment, "RPC component throws exception, the error code is {0}, the error message is: {1}", result, new Win32Exception((int)result).ToString());
            }

            this.site.Assert.AreEqual<uint>((uint)1, result, "NspiUnbind method should return 1 (Success).");
            return result;
        }

        /// <summary>
        /// The NspiQueryRows method returns a number of rows from a specified table to the client.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="tableCount">A DWORD value that contains the number values in the input parameter table. 
        /// This value is limited to 100,000.</param>
        /// <param name="table">An array of DWORD values, representing an Explicit Table.</param>
        /// <param name="count">A DWORD value that contains the number of rows the client is requesting.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value, 
        /// containing a list of the proptags of the properties that the client requires to be returned for each row returned.</param>
        /// <param name="rows">A nullable PropertyRowSet_r value, it contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiQueryRows(uint flags, ref STAT stat, uint tableCount, uint[] table, uint count, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result;
            IntPtr ptrRows = IntPtr.Zero;
            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);
            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiQueryRows(this.contextHandle, flags, ref stat, tableCount, table, count, ptrPropTags, out ptrRows);
                }
                catch (SEHException e)
                {
                    result = (int)NativeMethods.RpcExceptionCode(e);
                    this.site.Log.Add(LogEntryKind.Comment, "RPC component throws exception, the error code is {0}, the error message is: {1}", result, new Win32Exception(result).ToString());
                }

                if ((ErrorCodeValue)result == ErrorCodeValue.GeneralFailure && needRetry)
                {
                    Thread.Sleep(this.waitTime);
                }
                else
                {
                    break;
                }

                retryCount++;
            }
            while ((ErrorCodeValue)result == ErrorCodeValue.GeneralFailure && retryCount < this.maxRetryCount);

            if (!Enum.IsDefined(typeof(ErrorCodeValue), (uint)result))
            {
                throw new ArgumentException(string.Format("An unknown error is returned, the error code is: {0} and the error message is: {1}", result, new Win32Exception(result).ToString()));
            }

            if (propTags != null)
            {
                Marshal.FreeHGlobal(ptrPropTags);
            }

            // Parse rows according to ptrRows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRows);
            }

            // Free stat.
            Marshal.FreeHGlobal(ptrStat);
            return (ErrorCodeValue)result;
        }
        #endregion

        #region IDisposable Members
        /// <summary>
        /// This method is used to implement clean-up codes.
        /// </summary>
        /// <param name="disposing">Set TRUE to dispose resource otherwise set FALSE.</param>
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }
        #endregion
    }
}