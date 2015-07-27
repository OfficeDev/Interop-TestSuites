//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Storage Manifest data element 
    /// </summary>
    public class StorageManifestDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the StorageManifestDataElementData class.
        /// </summary>
        public StorageManifestDataElementData()
        {
            // Storage Manifest
            this.StorageManifestSchemaGUID = new StorageManifestSchemaGUID();
            this.StorageManifestRootDeclareList = new List<StorageManifestRootDeclare>();
        }

        /// <summary>
        /// Gets or sets storage manifest schema GUID.
        /// </summary>
        public StorageManifestSchemaGUID StorageManifestSchemaGUID { get; set; }

        /// <summary>
        /// Gets or sets storage manifest root declare.
        /// </summary>
        public List<StorageManifestRootDeclare> StorageManifestRootDeclareList { get; set; }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.StorageManifestSchemaGUID.SerializeToByteList());

            if (this.StorageManifestRootDeclareList != null)
            {
                foreach (StorageManifestRootDeclare storageManifestRootDeclare in this.StorageManifestRootDeclareList)
                {
                    byteList.AddRange(storageManifestRootDeclare.SerializeToByteList());
                }
            }

            return byteList;
        }

        /// <summary>
        /// Used to de-serialize data element.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The length of the array</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;

            this.StorageManifestSchemaGUID = StreamObject.GetCurrent<StorageManifestSchemaGUID>(byteArray, ref index);
            this.StorageManifestRootDeclareList = new List<StorageManifestRootDeclare>();

            StreamObjectHeaderStart header;
            int headerLength = 0;
            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                if (header.Type == StreamObjectTypeHeaderStart.StorageManifestRootDeclare)
                {
                    index += headerLength;
                    this.StorageManifestRootDeclareList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as StorageManifestRootDeclare);
                }
                else
                {
                    throw new DataElementParseErrorException(index, "Failed to parse StorageManifestDataElement, expect the inner object type StorageManifestRootDeclare, but actual type value is " + header.Type, null);
                }
            }

            return index - startIndex;
        }
    }

    /// <summary>
    /// Specifies one or more storage manifest root declare.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class StorageManifestRootDeclare : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the StorageManifestRootDeclare class.
        /// </summary>
        public StorageManifestRootDeclare()
            : base(StreamObjectTypeHeaderStart.StorageManifestRootDeclare)
        {
        }

        /// <summary>
        /// Gets or sets the root storage manifest.
        /// </summary>
        public ExGuid RootExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets the cell identifier.
        /// </summary>
        public CellID CellID { get; set; }

        /// <summary>
        /// Used to de-serialize the items.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.RootExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.CellID = BasicObject.Parse<CellID>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "StorageManifestRootDeclare", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The length of list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.RootExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.CellID.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Specifies a storage manifest schema GUID.
    /// </summary>
    public class StorageManifestSchemaGUID : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the StorageManifestSchemaGUID class.
        /// </summary>
        public StorageManifestSchemaGUID()
            : base(StreamObjectTypeHeaderStart.StorageManifestSchemaGUID)
        {
            // this.GUID = DataElementExGuids.StorageManifestGUID;
        }

        /// <summary>
        /// Gets or sets the schema GUID.
        /// </summary>
        public Guid GUID { get; set; }

        /// <summary>
        /// Used to de-serialize the items.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            byte[] temp = new byte[16];
            Array.Copy(byteArray, index, temp, 0, 16);
            this.GUID = new Guid(temp);
            index += 16;

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "StorageManifestSchemaGUID", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value 16</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            byteList.AddRange(this.GUID.ToByteArray());
            return 16;
        }
    }
}