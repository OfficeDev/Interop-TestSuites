//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Implementation for MS_OXCPRPTAdapter for ROPs call actions
    /// </summary>
    public partial class MS_OXCPRPTAdapter : ManagedAdapterBase, IMS_OXCPRPTAdapter
    {
        /// <summary>
        /// RopGetPropertiesSpecific implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertySizeLimit">This value specifies the maximum size allowed for a property value returned</param>
        /// <param name="wantUnicode">This value specifies whether to return string properties in Unicode</param>
        /// <param name="propertyTags">This field specifies the properties requested</param>
        /// <returns>Structure of RopGetPropertiesSpecificResponse</returns>
        private RopGetPropertiesSpecificResponse RopGetPropertiesSpecific(uint objHandle, ushort propertySizeLimit, ushort wantUnicode, PropertyTag[] propertyTags)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest;
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;

            getPropertiesSpecificRequest.RopId = (byte)RopId.RopGetPropertiesSpecific;
            getPropertiesSpecificRequest.LogonId = LogonId;
            getPropertiesSpecificRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            getPropertiesSpecificRequest.PropertySizeLimit = propertySizeLimit;
            getPropertiesSpecificRequest.WantUnicode = wantUnicode;
            if (propertyTags != null)
            {
                getPropertiesSpecificRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                // PropertyTags is null, so count of propertyTags is zero
                getPropertiesSpecificRequest.PropertyTagCount = 0x00;
            }

            getPropertiesSpecificRequest.PropertyTags = propertyTags;

            this.responseSOHsValue = this.ProcessSingleRop(getPropertiesSpecificRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.responseValue;
            return getPropertiesSpecificResponse;
        }

        /// <summary>
        /// RopGetPropertiesAllResponse implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertySizeLimit">This value specifies the maximum size allowed for a property value returned</param>
        /// <param name="wantUnicode">This value specifies whether to return string properties in Unicode</param>
        /// <returns>Structure of RopGetPropertiesAllResponse</returns>
        private RopGetPropertiesAllResponse RopGetPropertiesAll(uint objHandle, ushort propertySizeLimit, ushort wantUnicode)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetPropertiesAllRequest getPropertiesAllRequest;
            RopGetPropertiesAllResponse getPropertiesAllResponse;

            getPropertiesAllRequest.RopId = (byte)RopId.RopGetPropertiesAll;
            getPropertiesAllRequest.LogonId = LogonId;
            getPropertiesAllRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            getPropertiesAllRequest.PropertySizeLimit = propertySizeLimit;
            getPropertiesAllRequest.WantUnicode = wantUnicode;

            this.responseSOHsValue = this.ProcessSingleRop(getPropertiesAllRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.responseValue;
            return getPropertiesAllResponse;
        }

        /// <summary>
        /// RopGetPropertiesList implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <returns>Structure of RopGetPropertiesListResponse</returns>
        private RopGetPropertiesListResponse RopGetPropertiesList(uint objHandle)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetPropertiesListRequest getPropertiesListRequest;
            RopGetPropertiesListResponse getPropertiesListResponse;

            getPropertiesListRequest.RopId = (byte)RopId.RopGetPropertiesList;
            getPropertiesListRequest.LogonId = LogonId;
            getPropertiesListRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;

            this.responseSOHsValue = this.ProcessSingleRop(getPropertiesListRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getPropertiesListResponse = (RopGetPropertiesListResponse)this.responseValue;
            return getPropertiesListResponse;
        }

        /// <summary>
        /// RopSetProperties implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="taggedPropertyValueArray">Array of TaggedPropertyValue structures. This field specifies the property values to be set on the object</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>Structure of RopSetPropertiesResponse</returns>
        private RopSetPropertiesResponse RopSetProperties(uint objHandle, TaggedPropertyValue[] taggedPropertyValueArray, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSetPropertiesRequest setPropertiesRequest;
            RopSetPropertiesResponse setPropertiesResponse;

            setPropertiesRequest.RopId = (byte)RopId.RopSetProperties;
            setPropertiesRequest.LogonId = LogonId;
            setPropertiesRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            setPropertiesRequest.PropertyValues = taggedPropertyValueArray;
            if (setPropertiesRequest.PropertyValues != null)
            {
                setPropertiesRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;

                // Count is the size of setPropertiesRequest.PropertyValueCount.
                ushort count = sizeof(ushort);
                foreach (TaggedPropertyValue tagValue in taggedPropertyValueArray)
                {
                    count += (ushort)tagValue.Size();
                }

                setPropertiesRequest.PropertyValueSize = count;
            }
            else
            {
                // PropertyVaules is null, so count of PropertyValues is 0x00.
                setPropertiesRequest.PropertyValueCount = 0x00;

                // PropertyValueSize is the size of setPropertiesRequest.PropertyValueCount.
                setPropertiesRequest.PropertyValueSize = sizeof(ushort);
            }

            this.responseSOHsValue = this.ProcessSingleRop(setPropertiesRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            setPropertiesResponse = (RopSetPropertiesResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual(0x00, setPropertiesResponse.PropertyProblemCount, string.Format("RopSetProperties Failed! Error Count: 0x{0:X8}", setPropertiesResponse.PropertyProblemCount));
            }

            return setPropertiesResponse;
        }

        /// <summary>
        /// RopSetPropertiesNoReplicate implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="taggedPropertyValueArray">Array of TaggedPropertyValue structures. This field specifies the property values to be set on the object</param>
        /// <returns>Structure of  RopSetPropertiesNoReplicateResponse</returns>
        private RopSetPropertiesNoReplicateResponse RopSetPropertiesNoReplicate(uint objHandle, TaggedPropertyValue[] taggedPropertyValueArray)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSetPropertiesNoReplicateRequest setPropertiesNoReplicateRequest;
            RopSetPropertiesNoReplicateResponse setPropertiesNoReplicateResponse;

            setPropertiesNoReplicateRequest.RopId = (byte)RopId.RopSetPropertiesNoReplicate;
            setPropertiesNoReplicateRequest.LogonId = LogonId;
            setPropertiesNoReplicateRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            setPropertiesNoReplicateRequest.PropertyValues = taggedPropertyValueArray;
            if (setPropertiesNoReplicateRequest.PropertyValues != null)
            {
                // The count is the size of setPropertiesRequest.PropertyValueCount.
                ushort count = sizeof(ushort);
                foreach (TaggedPropertyValue tagValue in taggedPropertyValueArray)
                {
                    count += (ushort)tagValue.Size();
                }

                setPropertiesNoReplicateRequest.PropertyValueSize = count;
                setPropertiesNoReplicateRequest.PropertyValueCount = (ushort)taggedPropertyValueArray.Length;
            }
            else
            {
                // PropertyValues is null, so count of PropertyValues is 0x00
                setPropertiesNoReplicateRequest.PropertyValueCount = 0x00;

                // PropertyValueSize is the size of setPropertiesRequest.PropertyValueCount.
                setPropertiesNoReplicateRequest.PropertyValueSize = sizeof(ushort);
            }

            this.responseSOHsValue = this.ProcessSingleRop(setPropertiesNoReplicateRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            setPropertiesNoReplicateResponse = (RopSetPropertiesNoReplicateResponse)this.responseValue;
            return setPropertiesNoReplicateResponse;
        }

        /// <summary>
        /// RopDeleteProperties implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values to be deleted from the object.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>Structure of RopDeletePropertiesResponse</returns>
        private RopDeletePropertiesResponse RopDeleteProperties(uint objHandle, PropertyTag[] propertyTags, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopDeletePropertiesRequest deletePropertiesRequest;
            RopDeletePropertiesResponse deletePropertiesResponse;

            deletePropertiesRequest.RopId = (byte)RopId.RopDeleteProperties;
            deletePropertiesRequest.LogonId = LogonId;
            deletePropertiesRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            deletePropertiesRequest.PropertyTags = propertyTags;
            if (deletePropertiesRequest.PropertyTags != null)
            {
                deletePropertiesRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                // DeletePropertiesRequest.PropertyTags is null, so the length of deletePropertiesRequest.PropertyTags is 0x00.
                deletePropertiesRequest.PropertyTagCount = 0x00;
            }

            this.responseSOHsValue = this.ProcessSingleRop(deletePropertiesRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            deletePropertiesResponse = (RopDeletePropertiesResponse)this.responseValue;

            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, deletePropertiesResponse.ReturnValue, string.Format("RopDeleteProperties Failed! Error: 0x{0:X8}", deletePropertiesResponse.ReturnValue));
            }

            return deletePropertiesResponse;
        }

        /// <summary>
        /// RopDeletePropertiesNoReplicate implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the property values to be deleted from the object.</param>
        /// <returns>Structure of RopDeletePropertiesNoReplicateResponse</returns>
        private RopDeletePropertiesNoReplicateResponse RopDeletePropertiesNoReplicate(uint objHandle, PropertyTag[] propertyTags)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopDeletePropertiesNoReplicateRequest deletePropertiesNoReplicateRequest;
            RopDeletePropertiesNoReplicateResponse deletePropertiesNoReplicateResponse;

            deletePropertiesNoReplicateRequest.RopId = (byte)RopId.RopDeletePropertiesNoReplicate;
            deletePropertiesNoReplicateRequest.LogonId = LogonId;
            deletePropertiesNoReplicateRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            deletePropertiesNoReplicateRequest.PropertyTags = propertyTags;
            if (deletePropertiesNoReplicateRequest.PropertyTags != null)
            {
                deletePropertiesNoReplicateRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                // DeletePropertiesNoReplicateRequest.PropertyTags is null, so the count of deletePropertiesNoReplicateRequest.PropertyTags is 0x00.
                deletePropertiesNoReplicateRequest.PropertyTagCount = 0x00;
            }

            this.responseSOHsValue = this.ProcessSingleRop(deletePropertiesNoReplicateRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            deletePropertiesNoReplicateResponse = (RopDeletePropertiesNoReplicateResponse)this.responseValue;
            return deletePropertiesNoReplicateResponse;
        }

        /// <summary>
        /// The method is used to query an object for all the named properties. 
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored.</param>
        /// <param name="queryFlags">8-bit flags structure, These flags control how this ROP behaves</param>
        /// <param name="hasGuid">8-bit Boolean. This value specifies whether the PropertyGuid field is present.</param>
        /// <param name="propertyGuid">128-bit GUID. This field is present if HasGuid is nonzero and is not present if the value of the HasGuid field is zero. This value specifies the subset of named properties to be returned.</param>
        /// <returns>Structure of RopQueryNamedPropertiesResponse</returns>
        private RopQueryNamedPropertiesResponse RopQueryNamedProperties(uint objHandle, byte queryFlags, byte hasGuid, byte[] propertyGuid)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopQueryNamedPropertiesRequest queryNamedPropertiesRequest;
            RopQueryNamedPropertiesResponse queryNamedPropertiesResponse;

            queryNamedPropertiesRequest.RopId = (byte)RopId.RopQueryNamedProperties;
            queryNamedPropertiesRequest.LogonId = LogonId;
            queryNamedPropertiesRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            queryNamedPropertiesRequest.QueryFlags = queryFlags;
            queryNamedPropertiesRequest.HasGuid = hasGuid;
            queryNamedPropertiesRequest.PropertyGuid = propertyGuid;

            this.responseSOHsValue = this.ProcessSingleRop(queryNamedPropertiesRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            queryNamedPropertiesResponse = (RopQueryNamedPropertiesResponse)this.responseValue;
            return queryNamedPropertiesResponse;
        }

        /// <summary>
        /// RopCopyProperties implementation
        /// </summary>
        /// <param name="sourceHandle">This index specifies the location in the Server object handle table where the handle for source object.</param>
        /// <param name="destHandle">This index specifies the location in the Server object handle table where the handle for destination object</param>
        /// <param name="sourceHandleIndex">Unsigned 8-bit integer. This index specifies the location in the Server object handle table where the handle for the source Server object is stored</param>
        /// <param name="destHandleIndex">Unsigned 8-bit integer. This index specifies the location in the Server object handle table where the handle for the destination Server object is stored</param>
        /// <param name="wantAsynchronous">8-bit Boolean. This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress</param>
        /// <param name="copyFlags">8-bit flags structure. The possible values are specified in [MS-OXCPRPT]. These flags control the operation behavior</param>
        /// <param name="propertyTags">Array of PropertyTag structures. This field specifies the properties to copy</param>
        /// <returns>Response for RopCopyProperties, which could be a RopProgress response if server run asynchronously</returns>
        private object RopCopyProperties(uint sourceHandle, uint destHandle, byte sourceHandleIndex, byte destHandleIndex, byte wantAsynchronous, byte copyFlags, PropertyTag[] propertyTags)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCopyPropertiesRequest copyPropertiesRequest;

            copyPropertiesRequest.RopId = (byte)RopId.RopCopyProperties;
            copyPropertiesRequest.LogonId = LogonId;
            copyPropertiesRequest.SourceHandleIndex = sourceHandleIndex;
            copyPropertiesRequest.DestHandleIndex = destHandleIndex;
            copyPropertiesRequest.WantAsynchronous = wantAsynchronous;
            copyPropertiesRequest.CopyFlags = copyFlags;
            copyPropertiesRequest.PropertyTags = propertyTags;
            if (copyPropertiesRequest.PropertyTags != null)
            {
                copyPropertiesRequest.PropertyTagCount = (ushort)propertyTags.Length;
            }
            else
            {
                // CopyPropertiesRequest.PropertyTags is null, so the count of copyPropertiesRequest.PropertyTags is 0x00.
                copyPropertiesRequest.PropertyTagCount = 0x00;
            }

            List<uint> handleList = new List<uint>
            {
                sourceHandle, destHandle
            };

            this.responseSOHsValue = this.ProcessSingleRopWithMutipleServerObjects(copyPropertiesRequest, handleList, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            return this.responseValue;
        }

        /// <summary>
        /// RopProgress implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored</param>
        /// <param name="wantCancel">8-bit Boolean. This value specifies whether to cancel the operation</param>
        /// <returns>Response of RopProgress, can be other ROPs' response if server run asynchronously</returns>
        private object RopProgress(uint objHandle, byte wantCancel)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopProgressRequest progressRequest;
            progressRequest.RopId = (byte)RopId.RopProgress;
            progressRequest.LogonId = LogonId;
            progressRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            progressRequest.WantCancel = wantCancel;
            this.responseSOHsValue = this.ProcessSingleRop(progressRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            return this.responseValue;
        }

        /// <summary>
        /// RopCopyTo implementation
        /// </summary>
        /// <param name="sourceHandle">This index specifies the location in the Server object handle table where the handle for source object.</param>
        /// <param name="destHandle">This index specifies the location in the Server object handle table where the handle for destination object</param>
        /// <param name="sourceHandleIndex">Unsigned 8-bit integer. This index specifies the location in the Server object handle table where the handle for the source Server object is stored</param>
        /// <param name="destHandleIndex">Unsigned 8-bit integer. This index specifies the location in the Server object handle table where the handle for the destination Server object is stored</param>
        /// <param name="wantAsynchronous">8-bit Boolean. This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress</param>
        /// <param name="wantSubObjects">8-bit Boolean. This value specifies whether to copy sub-objects</param>
        /// <param name="copyFlags">8-bit flags structure. The possible values are specified in [MS-OXCPRPT]. These flags control the operation behavior</param>
        /// <param name="excludedTags">Array of PropertyTag structures. This field specifies the properties to exclude from the copy</param>
        /// <returns>Response for RopCopyTo, which could be a RopProgress response if server run asynchronously</returns>
        private object RopCopyTo(uint sourceHandle, uint destHandle, byte sourceHandleIndex, byte destHandleIndex, byte wantAsynchronous, byte wantSubObjects, byte copyFlags, PropertyTag[] excludedTags)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCopyToRequest copyToRequest;

            copyToRequest.RopId = (byte)RopId.RopCopyTo;
            copyToRequest.LogonId = LogonId;
            copyToRequest.SourceHandleIndex = sourceHandleIndex;
            copyToRequest.DestHandleIndex = destHandleIndex;
            copyToRequest.WantAsynchronous = wantAsynchronous;
            copyToRequest.WantSubObjects = wantSubObjects;
            copyToRequest.CopyFlags = copyFlags;
            copyToRequest.ExcludedTags = excludedTags;
            if (copyToRequest.ExcludedTags != null)
            {
                copyToRequest.ExcludedTagCount = (ushort)excludedTags.Length;
            }
            else
            {
                // CopyToRequest.ExcludedTags is null, so the count of copyToRequest.ExcludedTags is 0x00.
                copyToRequest.ExcludedTagCount = 0x00;
            }

            List<uint> handleList = new List<uint>
            {
                sourceHandle, destHandle
            };

            this.responseSOHsValue = this.ProcessSingleRopWithMutipleServerObjects(copyToRequest, handleList, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            return this.responseValue;
        }

        /// <summary>
        /// RopGetPropertyIdsFromNames implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored</param>
        /// <param name="flags">8-bit flags structure. These flags control the behavior of this operation</param>
        /// <param name="propertyNames">List of PropertyName structures. This field specifies the property names requested.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>Structure of RopGetPropertyIdsFromNamesResponse</returns>
        private RopGetPropertyIdsFromNamesResponse RopGetPropertyIdsFromNames(uint objHandle, byte flags, PropertyName[] propertyNames, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetPropertyIdsFromNamesRequest getPropertyIdsFromNamesRequest;
            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse;

            getPropertyIdsFromNamesRequest.RopId = (byte)RopId.RopGetPropertyIdsFromNames;
            getPropertyIdsFromNamesRequest.LogonId = LogonId;
            getPropertyIdsFromNamesRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            getPropertyIdsFromNamesRequest.Flags = flags;
            getPropertyIdsFromNamesRequest.PropertyNames = propertyNames;
            if (getPropertyIdsFromNamesRequest.PropertyNames != null)
            {
                getPropertyIdsFromNamesRequest.PropertyNameCount = (ushort)propertyNames.Length;
            }
            else
            {
                // GetPropertyIdsFromNamesRequest.PropertyNames is null, so getPropertyIdsFromNamesRequest.PropertyNames is 0x00.
                getPropertyIdsFromNamesRequest.PropertyNameCount = 0x00;
            }

            this.responseSOHsValue = this.ProcessSingleRop(getPropertyIdsFromNamesRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getPropertyIdsFromNamesResponse = (RopGetPropertyIdsFromNamesResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, getPropertyIdsFromNamesResponse.ReturnValue, string.Format("RopGetPropertyIdsFromNames Failed! Error: 0x{0:X8}", getPropertyIdsFromNamesResponse.ReturnValue));
            }

            return getPropertyIdsFromNamesResponse;
        }

        /// <summary>
        /// RopGetNamesFromPropertyIds implementation
        /// </summary>
        /// <param name="objHandle">This index specifies the location in the Server object handle table where the handle for the input Server object is stored</param>
        /// <param name="propertyIds">Array of unsigned 16-bit integers, number of integers in the array is specified by the PropertyIdCount field.</param>
        /// <returns>Structure of RopGetNamesFromPropertyIdsResponse</returns>
        private RopGetNamesFromPropertyIdsResponse RopGetNamesFromPropertyIds(uint objHandle, PropertyId[] propertyIds)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopGetNamesFromPropertyIdsRequest getNamesFromPropertyIdsRequest;
            RopGetNamesFromPropertyIdsResponse getNamesFromPropertyIdsResponse;

            getNamesFromPropertyIdsRequest.RopId = (byte)RopId.RopGetNamesFromPropertyIds;
            getNamesFromPropertyIdsRequest.LogonId = LogonId;
            getNamesFromPropertyIdsRequest.InputHandleIndex = (byte)HandleIndex.FirstIndex;
            getNamesFromPropertyIdsRequest.PropertyIds = propertyIds;
            if (getNamesFromPropertyIdsRequest.PropertyIds != null)
            {
                getNamesFromPropertyIdsRequest.PropertyIdCount = (ushort)propertyIds.Length;
            }
            else
            {
                // GetNamesFromPropertyIdsRequest.PropertyIds is null, so count of getNamesFromPropertyIdsRequest.PropertyIds is 0x00.
                getNamesFromPropertyIdsRequest.PropertyIdCount = 0x00;
            }

            this.responseSOHsValue = this.ProcessSingleRop(getNamesFromPropertyIdsRequest, objHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            getNamesFromPropertyIdsResponse = (RopGetNamesFromPropertyIdsResponse)this.responseValue;
            return getNamesFromPropertyIdsResponse;
        }
    }
}