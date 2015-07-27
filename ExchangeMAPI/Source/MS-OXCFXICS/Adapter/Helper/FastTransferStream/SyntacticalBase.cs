//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;

    /// <summary>
    /// Base class for all syntactical object.
    /// </summary>
    public abstract class SyntacticalBase : IStreamSerializable, IStreamDeserializable
    {
        /// <summary>
        /// The size of an PidTag value.
        /// </summary>
        protected const int PidLength = MarkersHelper.PidTagLength;

        #region Param
        /// <summary>
        /// Previous position
        /// </summary>
        private long previousPosition;

        /// <summary>
        /// A list of errorInfo objects.
        /// </summary>
        private List<ErrorInfo> errorInfoList;

        /// <summary>
        /// Indicate whether a stream MUST NOT be split within a single atom.
        /// </summary>
        private bool isNotSplitedInSingleItem;

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the SyntacticalBase class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        protected SyntacticalBase(FastTransferStream stream)
        {
            this.previousPosition = stream.Position;
            if (stream != null && stream.Length > 0)
            {
                this.errorInfoList = new List<ErrorInfo>();
                this.Deserialize(stream);
            }

            // No exception set flag.
            this.isNotSplitedInSingleItem = true;
        }
        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the allPropList.
        /// </summary>
        public static List<PropList> AllPropList
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets a value indicating whether a stream MUST NOT be split within a single atom.
        /// </summary>
        public bool IsNotSplitedInSingleItem
        {
            get { return this.isNotSplitedInSingleItem; }
            set { this.isNotSplitedInSingleItem = value; }
        }

        /// <summary>
        /// Gets the errorInfo list.
        /// </summary>
        public List<ErrorInfo> ErrorInfoList
        {
            get { return this.errorInfoList; }
        }

        /// <summary>
        /// Gets a value indicating whether has errorInfo.
        /// </summary>
        public bool HasErrorInfo
        {
            get
            {
                return this.errorInfoList.Count > 0;
            }
        }

        /// <summary>
        /// Gets previous position.
        /// </summary>
        protected long PreviousPosition
        {
            get
            {
                return this.previousPosition;
            }
        }
        #endregion

        #endregion

        #region Interfaces
        /// <summary>
        /// Serialize object to a FastTransferStream.
        /// </summary>
        /// <returns>A FastTransferStream contains the serialized object.</returns>
        public virtual FastTransferStream Serialize()
        {
            AdapterHelper.Site.Assert.Fail("Method is not implemented.");
            return null;
        }

        /// <summary>
        /// Deserialize object from memory stream,
        /// after deserialization stream's read position += serialized object size;
        /// </summary>
        /// <param name="stream">Stream contains the serialized object</param>
        public abstract void Deserialize(FastTransferStream stream);

        /// <summary>
        /// Deserialize a nested structure defined as StartMarker content EndMarker
        /// </summary>
        /// <typeparam name="T">The deserialized type</typeparam>
        /// <param name="stream">A FastTransferStream</param>
        /// <param name="startMarker">The start marker of the nested structure</param>
        /// <param name="endMarker">The end marker of the nested structure</param>
        /// <param name="member">The deserialized nested structure</param>
        public void Deserialize<T>(
            FastTransferStream stream,
            Markers startMarker,
            Markers endMarker,
            out T member) where T : SyntacticalBase
        {
            Type subType = typeof(T);
            this.FindVerify(subType);

            if (stream.ReadMarker(startMarker))
            {
                object tmp = subType.Assembly.CreateInstance(
                    subType.FullName,
                    false,
                    BindingFlags.CreateInstance,
                    null,
                    new object[] { stream },
                    null,
                    null);
                if (stream.ReadMarker(endMarker))
                {
                    member = tmp as T;
                    Debug.Assert(member != null, "The deserialization operation should be successful.");
                    return;
                }
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
            member = null;
        }

        /// <summary>
        /// Deserialize structure defined as StartMarker content
        /// </summary>
        /// <typeparam name="T">The type of the deserialized structure</typeparam>
        /// <param name="stream">A FastTransferStream</param>
        /// <param name="startMarker">The start marker of the deserialized structure</param>
        /// <param name="member">The deserialized structure</param>
        public void Deserialize<T>(
            FastTransferStream stream,
            Markers startMarker,
            out T member) where T : SyntacticalBase
        {
            Type subType = typeof(T);
            this.FindVerify(subType);

            if (stream.ReadMarker(startMarker))
            {
                object tmp = subType.Assembly.CreateInstance(
                    subType.FullName,
                    false,
                    BindingFlags.CreateInstance,
                    null,
                    new object[] { stream },
                    null,
                    null);
                member = tmp as T;
                AdapterHelper.Site.Assert.IsNotNull(member, "The deserialization operation should be successful.");
                return;
            }

            AdapterHelper.Site.Assert.Fail("The stream cannot be deserialized successfully.");
            member = null;
        }

        /// <summary>
        /// Check errorInfos in a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream</param>
        /// <param name="offset">An offset to the current position of the FastTransferStream</param>
        /// <returns>If the FastTransferStream contains errorInfos then read them and
        /// return true, else return false.
        /// </returns>
        protected bool CheckErrorInfo(FastTransferStream stream, int offset)
        {
            int count = 0;
            stream.Position += offset;
            while (stream.VerifyErrorInfo(0))
            {
                this.errorInfoList.Add(new ErrorInfo(stream));
                count++;
            }

            return count > 0;
        }

        /// <summary>
        /// Pop a FastTransferStream's position.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        protected void Pop(FastTransferStream stream)
        {
            stream.Position = this.previousPosition;
        }

        /// <summary>
        /// Find the Verify method.
        /// </summary>
        /// <param name="type">The Type which contains the method.</param>
        private void FindVerify(Type type)
        {
            bool hasType = false;
            MethodInfo[] methods = type.GetMethods(BindingFlags.Static
                | BindingFlags.Public);
            foreach (MethodInfo mi in methods)
            {
                if (mi.Name == "Verify")
                {
                    hasType = true;
                }
            }

            AdapterHelper.Site.Assert.IsTrue(hasType, "The type {0} should exist.", type.Name);
        }
        #endregion
    }
}