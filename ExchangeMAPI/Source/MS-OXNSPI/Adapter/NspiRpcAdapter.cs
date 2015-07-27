//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
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
        /// The NspiGetSpecialTable method returns the rows of a special table to the client. 
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A pointer to a STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="version">A reference to a DWORD. On input, it holds the value of the version number of
        /// the address book hierarchy table that the client has. On output, it holds the version of the server's address book hierarchy table.</param>
        /// <param name="rows">A PropertyRowSet_r structure. On return, it holds the rows for the table that the client is requesting.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetSpecialTable(uint flags, ref STAT stat, ref uint version, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result;
            IntPtr ptrRows = IntPtr.Zero;
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiGetSpecialTable(this.contextHandle, flags, ref stat, ref version, out ptrRows);
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

            // Parse PropertyRowSet_r from ptrRows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRows);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiUpdateStat method updates the STAT block that represents the position in a table 
        /// to reflect positioning changes requested by the client.
        /// </summary>
        /// <param name="reserved">A DWORD value. Reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A pointer to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="delta">The value NULL or a pointer to a LONG value that indicates movement 
        /// within the address book container specified by the input parameter stat.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiUpdateStat(uint reserved, ref STAT stat, ref int? delta, bool needRetry = true)
        {
            int result;
            IntPtr ptrDelta = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);

            if (delta != null)
            {
                ptrDelta = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
                Marshal.WriteInt32(ptrDelta, delta.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiUpdateStat(this.contextHandle, reserved, ref stat, ptrDelta);
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

            // Parse delta.
            if (ptrDelta == IntPtr.Zero)
            {
                delta = null;
            }
            else
            {
                delta = Marshal.ReadInt32(ptrDelta);

                // Free ptrDelta pointing memory.
                Marshal.FreeHGlobal(ptrDelta);
            }

            // Free stat pointing memory.
            Marshal.FreeHGlobal(ptrStat);

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiQueryColumns method returns a list of all the properties that the server is aware of. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="columns">A PropertyTagArray_r structure that contains a list of proptags.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiQueryColumns(uint reserved, uint flags, out PropertyTagArray_r? columns, bool needRetry = true)
        {
            int result;
            IntPtr ptrColumns = IntPtr.Zero;
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiQueryColumns(this.contextHandle, reserved, flags, out ptrColumns);
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

            if (ptrColumns == IntPtr.Zero)
            {
                columns = null;
            }
            else
            {
                columns = AdapterHelper.ParsePropertyTagArray_r(ptrColumns);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiGetPropList method returns a list of all the properties that have values on a specified object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="mid">A DWORD value that contains a Minimal Entry ID.</param>
        /// <param name="codePage">The code page in which the client wants the server to express string values properties.</param>
        /// <param name="propTags">A PropertyTagArray_r value. On return, it holds a list of properties.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetPropList(uint flags, uint mid, uint codePage, out PropertyTagArray_r? propTags, bool needRetry = true)
        {
            int result;
            IntPtr ptrPropTags = IntPtr.Zero;
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiGetPropList(this.contextHandle, flags, mid, codePage, out ptrPropTags);
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

            // Parse propTags.
            if (ptrPropTags == IntPtr.Zero)
            {
                propTags = null;
            }
            else
            {
                propTags = AdapterHelper.ParsePropertyTagArray_r(ptrPropTags);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiGetProps method returns an address book row that contains a set of the properties
        /// and values that exist on an object.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the properties that the client wants to be returned.</param>
        /// <param name="rows">A nullable PropertyRow_r value. 
        /// It contains the address book container row the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetProps(uint flags, STAT stat, PropertyTagArray_r? propTags, out PropertyRow_r? rows, bool needRetry = true)
        {
            int result;
            IntPtr ptrRows = IntPtr.Zero;
            IntPtr ptrPropTags = IntPtr.Zero;
            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiGetProps(this.contextHandle, flags, ref stat, ptrPropTags, out ptrRows);
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

            // Parse rows from ptrRows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRow_r(ptrRows);
            }

            return (ErrorCodeValue)result;
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

        /// <summary>
        /// The NspiSeekEntries method searches for and sets the logical position in a specific table
        /// to the first entry greater than or equal to a specified value. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="target">A PropertyValue_r value holding the value that is being sought.</param>
        /// <param name="table">The value NULL or a PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">It contains a list of the proptags of the columns 
        /// that client wants to be returned for each row returned.</param>
        /// <param name="rows">It contains the address book container rows the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiSeekEntries(uint reserved, ref STAT stat, PropertyValue_r target, PropertyTagArray_r? table, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result = 0;
            IntPtr ptrRows = IntPtr.Zero;
            IntPtr ptrETable = IntPtr.Zero;
            IntPtr ptrTarget = AdapterHelper.AllocPropertyValue_r(target);
            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);
            if (table != null)
            {
                ptrETable = AdapterHelper.AllocPropertyTagArray_r(table.Value);
            }

            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiSeekEntries(this.contextHandle, reserved, ref stat, ptrTarget, ptrETable, ptrPropTags, out ptrRows);
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

            if (table != null)
            {
                // Free ptrETable.
                Marshal.FreeHGlobal(ptrETable);
            }

            if (propTags != null)
            {
                // Free ptrPropTags.
                Marshal.FreeHGlobal(ptrPropTags);
            }

            // Free ptrTarget.
            AdapterHelper.FreePropertyValue_r(ptrTarget);

            // Free stat.
            Marshal.FreeHGlobal(ptrStat);

            // Parse rows from ptrRows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRows);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiGetMatches method returns an Explicit Table. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block describing a logical position in a specific address book container.</param>
        /// <param name="proReserved">A PropertyTagArray_r reserved for future use.</param>
        /// <param name="reserved2">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="filter">The value NULL or a Restriction_r value. 
        /// It holds a logical restriction to apply to the rows in the address book container specified in the stat parameter.</param>
        /// <param name="propName">The value NULL or a PropertyName_r value. 
        /// It holds the property to be opened as a restricted address book container.</param>
        /// <param name="requested">A DWORD value. It contains the maximum number of rows to return in a restricted address book container.</param>
        /// <param name="outMids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value. 
        /// It contains a list of the proptags of the columns that client wants to be returned for each row returned.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. It contains the address book container rows the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetMatches(uint reserved, ref STAT stat, PropertyTagArray_r? proReserved, uint reserved2, Restriction_r? filter, PropertyName_r? propName, uint requested, out PropertyTagArray_r? outMids, PropertyTagArray_r? propTags, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result = 0;
            Restriction_r[] filters = null;
            IntPtr ptrRows = IntPtr.Zero;
            IntPtr ptrpReserved = IntPtr.Zero;
            IntPtr ptrFilter = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);
            IntPtr ptrLpPropName = IntPtr.Zero;
            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptrOutMids = IntPtr.Zero;
            if (proReserved != null)
            {
                ptrpReserved = AdapterHelper.AllocPropertyTagArray_r(proReserved.Value);
            }

            if (filter != null)
            {
                filters = new Restriction_r[1];
                filters[0] = filter.Value;

                ptrFilter = AdapterHelper.AllocRestriction_rs(filters);
            }

            if (propName != null)
            {
                ptrLpPropName = AdapterHelper.AllocPropertyName_r(propName.Value);
            }

            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiGetMatches(this.contextHandle, reserved, ref stat, ptrpReserved, reserved2, ptrFilter, ptrLpPropName, requested, out ptrOutMids, ptrPropTags, out ptrRows);
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

            if (proReserved != null)
            {
                Marshal.FreeHGlobal(ptrpReserved);
            }

            if (filter != null)
            {
                AdapterHelper.FreeRestriction_rs(ptrFilter, filters.Length);
            }

            if (propName != null)
            {
                AdapterHelper.FreePropertyName_r(ptrLpPropName);
            }

            if (propTags != null)
            {
                Marshal.FreeHGlobal(ptrPropTags);
            }

            // Free stat.
            Marshal.FreeHGlobal(ptrStat);

            // Parse outmids.
            if (ptrOutMids == IntPtr.Zero)
            {
                outMids = null;
            }
            else
            {
                outMids = AdapterHelper.ParsePropertyTagArray_r(ptrOutMids);
            }

            // Parse rows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRows);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiResortRestriction method applies to a sort order to the objects in a restricted address book container.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A reference to a STAT block describing a logical position in a specific address book container.</param>
        /// <param name="proInMIds">A PropertyTagArray_r value. 
        /// It holds a list of Minimal Entry IDs that comprise a restricted address book container.</param>
        /// <param name="outMIds">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs 
        /// that comprise a restricted address book container.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResortRestriction(uint reserved, ref STAT stat, PropertyTagArray_r proInMIds, ref PropertyTagArray_r? outMIds, bool needRetry = true)
        {
            int result;
            IntPtr ptrInMIds = AdapterHelper.AllocPropertyTagArray_r(proInMIds);
            IntPtr ptrOutMIds = IntPtr.Zero;
            IntPtr tempPtr = IntPtr.Zero;
            IntPtr ptrStat = AdapterHelper.AllocStat(stat);
            if (outMIds != null)
            {
                ptrOutMIds = AdapterHelper.AllocPropertyTagArray_r(outMIds.Value);
                tempPtr = new IntPtr(ptrOutMIds.ToInt32());
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiResortRestriction(this.contextHandle, reserved, ref stat, ptrInMIds, ref ptrOutMIds);
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

            // Parse mids.
            if (ptrOutMIds == IntPtr.Zero)
            {
                outMIds = null;
            }
            else
            {
                outMIds = AdapterHelper.ParsePropertyTagArray_r(ptrOutMIds);
            }

            // Free ptrOutMIds.
            if (tempPtr != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(tempPtr);
            }

            // Free stat.
            Marshal.FreeHGlobal(ptrStat);

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiCompareMIds method compares the position in an address book container of two objects 
        /// identified by Minimal Entry ID and returns the value of the comparison.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="mid1">The mid1 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="mid2">The mid2 is a DWORD value containing a Minimal Entry ID.</param>
        /// <param name="results">A DWORD value. On return, it contains the result of the comparison.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiCompareMIds(uint reserved, STAT stat, uint mid1, uint mid2, out int results, bool needRetry = true)
        {
            int result;
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiCompareMIds(this.contextHandle, reserved, ref stat, mid1, mid2, out results);
                }
                catch (SEHException e)
                {
                    result = (int)NativeMethods.RpcExceptionCode(e);
                    results = 0;
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

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiDNToMId method maps a set of DN to a set of Minimal Entry ID.
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use. Ignored by the server.</param>
        /// <param name="names">A StringsArray_r value. It holds a list of strings that contain DNs.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it holds a list of Minimal Entry IDs.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiDNToMId(uint reserved, StringsArray_r names, out PropertyTagArray_r? mids, bool needRetry = true)
        {
            int result;

            IntPtr ptrMIds = IntPtr.Zero;
            IntPtr ptrNames = AdapterHelper.AllocStringsArray_r(names);
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiDNToMId(this.contextHandle, reserved, ptrNames, out ptrMIds);
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

            // Free ptrNames.
            AdapterHelper.FreeStringsArray_r(ptrNames);

            // Parse mids.
            if (ptrMIds == IntPtr.Zero)
            {
                mids = null;
            }
            else
            {
                mids = AdapterHelper.ParsePropertyTagArray_r(ptrMIds);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiModProps method is used to modify the properties of an object in the address book. 
        /// </summary>
        /// <param name="reserved">A DWORD value reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r. 
        /// It contains a list of the proptags of the columns from which the client requests all the values to be removed.</param>
        /// <param name="row">A PropertyRow_r value. It contains an address book row.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiModProps(uint reserved, STAT stat, PropertyTagArray_r? propTags, PropertyRow_r row, bool needRetry = true)
        {
            int result;

            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptrRow = AdapterHelper.AllocPropertyRow_r(row);
            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiModProps(this.contextHandle, reserved, ref stat, ptrPropTags, ptrRow);
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

            AdapterHelper.FreePropertyRow_r(ptrRow);

            if (propTags != null)
            {
                Marshal.FreeHGlobal(ptrPropTags);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiModLinkAtt method modifies the values of a specific property of a specific row in the address book.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="propTag">A DWORD value. It contains the proptag of the property that the client wants to modify.</param>
        /// <param name="mid">A DWORD value that contains the Minimal Entry ID of the address book row that the client wants to modify.</param>
        /// <param name="entryIds">A BinaryArray value. It contains a list of EntryIDs to be used to modify the requested property on the requested address book row.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiModLinkAtt(uint flags, uint propTag, uint mid, BinaryArray_r entryIds, bool needRetry = true)
        {
            int result;
            IntPtr ptrEntryIds = AdapterHelper.AllocBinaryArray_r(entryIds);
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiModLinkAtt(this.contextHandle, flags, propTag, mid, ptrEntryIds);
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

            AdapterHelper.FreeBinaryArray_r(ptrEntryIds);
            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiResolveNames method takes a set of string values in an 8-bit character set 
        /// and performs ANR on those strings. 
        /// </summary>
        /// <param name="reserved">A DWORD reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r value containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="stringArray">A StringsArray_r value. It specifies the values on which the client is requesting the server to do ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r value. 
        /// It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResolveNames(uint reserved, STAT stat, PropertyTagArray_r? propTags, StringsArray_r? stringArray, out PropertyTagArray_r? mids, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result;

            IntPtr ptrRows = IntPtr.Zero;
            STAT[] stats = new STAT[1];
            stats[0] = stat;
            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptraStr = IntPtr.Zero;
            IntPtr ptrMIds = IntPtr.Zero;

            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            if (stringArray != null)
            {
                ptraStr = AdapterHelper.AllocStringsArray_r(stringArray.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiResolveNames(this.contextHandle, reserved, stats, ptrPropTags, ptraStr, out ptrMIds, out ptrRows);
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

            if (stringArray != null)
            {
                AdapterHelper.FreeStringsArray_r(ptraStr);
            }

            // Parse mids.
            if (ptrMIds == IntPtr.Zero)
            {
                mids = null;
            }
            else
            {
                mids = AdapterHelper.ParsePropertyTagArray_r(ptrMIds);
            }

            // Parse ppRows.
            if (ptrRows == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRows);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiResolveNamesW method takes a set of string values in the Unicode character set 
        /// and performs ANR on those strings. 
        /// </summary>
        /// <param name="reserved">A DWORD value that is reserved for future use.</param>
        /// <param name="stat">A STAT block that describes a logical position in a specific address book container.</param>
        /// <param name="propTags">The value NULL or a reference to a PropertyTagArray_r containing a list of the proptags of the columns 
        /// that the client requests to be returned for each row returned.</param>
        /// <param name="wstr">A WStringsArray_r value. It specifies the values on which the client is requesting the server to perform ANR.</param>
        /// <param name="mids">A PropertyTagArray_r value. On return, it contains a list of Minimal Entry IDs that match the array of strings.</param>
        /// <param name="rows">A reference to a PropertyRowSet_r structure. It contains the address book container rows that the server returns in response to the request.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiResolveNamesW(uint reserved, STAT stat, PropertyTagArray_r? propTags, WStringsArray_r? wstr, out PropertyTagArray_r? mids, out PropertyRowSet_r? rows, bool needRetry = true)
        {
            int result;

            IntPtr ptrRow = IntPtr.Zero;
            STAT[] stats = new STAT[1];
            stats[0] = stat;
            IntPtr ptrPropTags = IntPtr.Zero;
            IntPtr ptraWStr = IntPtr.Zero;
            IntPtr ptrMIds = IntPtr.Zero;

            if (propTags != null)
            {
                ptrPropTags = AdapterHelper.AllocPropertyTagArray_r(propTags.Value);
            }

            if (wstr != null)
            {
                ptraWStr = AdapterHelper.AllocWStringsArray_r(wstr.Value);
            }

            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiResolveNamesW(this.contextHandle, reserved, stats, ptrPropTags, ptraWStr, out ptrMIds, out ptrRow);
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

            // Free memory ptrPropTags pointing.
            if (propTags != null)
            {
                Marshal.FreeHGlobal(ptrPropTags);
            }

            // Free memory ptraWStr pointing.
            if (wstr != null)
            {
                AdapterHelper.FreeStringsArray_r(ptraWStr);
            }

            // Parse mids.
            if (ptrMIds == IntPtr.Zero)
            {
                mids = null;
            }
            else
            {
                mids = AdapterHelper.ParsePropertyTagArray_r(ptrMIds);
            }

            // Parse ppRows.
            if (ptrRow == IntPtr.Zero)
            {
                rows = null;
            }
            else
            {
                rows = AdapterHelper.ParsePropertyRowSet_r(ptrRow);
            }

            return (ErrorCodeValue)result;
        }

        /// <summary>
        /// The NspiGetTemplateInfo method returns information about template objects.
        /// </summary>
        /// <param name="flags">A DWORD value that contains a set of bit flags.</param>
        /// <param name="type">A DWORD value. It specifies the display type of the template for which the information is requested.</param>
        /// <param name="dn">The value NULL or the DN of the template requested. The value is NULL-terminated.</param>
        /// <param name="codePage">A DWORD value. It specifies the code page of the template for which the information is requested.</param>
        /// <param name="localeID">A DWORD value. It specifies the LCID of the template for which the information is requested.</param>
        /// <param name="data">A reference to a PropertyRow_r value. On return, it contains the information requested.</param>
        /// <param name="needRetry">A bool value indicates if need to retry to get an expected result. This parameter is designed to avoid meaningless retry when an error response is expected.</param>
        /// <returns>Status of NSPI method.</returns>
        public ErrorCodeValue NspiGetTemplateInfo(uint flags, uint type, string dn, uint codePage, uint localeID, out PropertyRow_r? data, bool needRetry = true)
        {
            int result;

            IntPtr ptrData = IntPtr.Zero;
            int retryCount = 0;
            do
            {
                try
                {
                    result = OxnspiInterop.NspiGetTemplateInfo(this.contextHandle, flags, type, dn, codePage, localeID, out ptrData);
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

            // Parse data.
            if (ptrData == IntPtr.Zero)
            {
                data = null;
            }
            else
            {
                data = AdapterHelper.ParsePropertyRow_r(ptrData);
            }

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