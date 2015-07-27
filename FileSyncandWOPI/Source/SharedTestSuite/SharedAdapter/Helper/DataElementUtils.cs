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
    using System.Linq;

    /// <summary>
    /// This class is used to build data element list or parse the data element list.
    /// </summary>
    public static class DataElementUtils
    {
        /// <summary>
        /// The constant value of root extended GUID.
        /// </summary>
        public static readonly Guid RootExGuid = new Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073");
        
        /// <summary>
        /// The constant value of cell second extended GUID.
        /// </summary>
        public static readonly Guid CellSecondExGuid = new Guid("6F2A4665-42C8-46C7-BAB4-E28FDCE1E32B");
        
        /// <summary>
        /// The constant value of schema GUID.
        /// </summary>
        public static readonly Guid SchemaGuid = new Guid("0EB93394-571D-41E9-AAD3-880D92D31955");

        /// <summary>
        /// This method is used to build a list of data elements to represent a file.
        /// </summary>
        /// <param name="fileContent">Specify the file content byte array.</param>
        /// <param name="storageIndexExGuid">Output parameter to represent the storage index GUID.</param>
        /// <returns>Return the list of data elements.</returns>
        public static List<DataElement> BuildDataElements(byte[] fileContent, out ExGuid storageIndexExGuid)
        {
            List<DataElement> dataElementList = new List<DataElement>();
            ExGuid rootNodeObjectExGuid;
            List<ExGuid> objectDataExGuidList = new List<ExGuid>();
            dataElementList.AddRange(CreateObjectGroupDataElement(fileContent, out rootNodeObjectExGuid, ref objectDataExGuidList));

            ExGuid baseRevisionID = new ExGuid(0u, Guid.Empty);
            Dictionary<ExGuid, ExGuid> revisionMapping = new Dictionary<ExGuid, ExGuid>();
            ExGuid currentRevisionID;
            dataElementList.Add(CreateRevisionManifestDataElement(rootNodeObjectExGuid, baseRevisionID, objectDataExGuidList, ref revisionMapping, out currentRevisionID));

            Dictionary<CellID, ExGuid> cellIDMapping = new Dictionary<CellID, ExGuid>();
            dataElementList.Add(CreateCellMainifestDataElement(currentRevisionID, ref cellIDMapping));

            dataElementList.Add(CreateStorageManifestDataElement(cellIDMapping));
            dataElementList.Add(CreateStorageIndexDataElement(dataElementList.Last().DataElementExtendedGUID, cellIDMapping, revisionMapping));

            storageIndexExGuid = dataElementList.Last().DataElementExtendedGUID;
            return dataElementList;
        }

        /// <summary>
        /// This method is used to create object group data/blob element list.
        /// </summary>
        /// <param name="fileContent">Specify the file content in byte array format.</param>
        /// <param name="rootNodeExGuid">Output parameter to represent the root node extended GUID.</param>
        /// <param name="objectDataExGuidList">Input/Output parameter to represent the list of extended GUID for the data object data.</param>
        /// <returns>Return the list of data element which will represent the file content.</returns>
        public static List<DataElement> CreateObjectGroupDataElement(byte[] fileContent, out ExGuid rootNodeExGuid, ref List<ExGuid> objectDataExGuidList)
        {
            NodeObject rootNode = new RootNodeObject.RootNodeObjectBuilder().Build(fileContent);
            
            // Storage the root object node ExGuid
            rootNodeExGuid = new ExGuid(rootNode.ExGuid);
            List<DataElement> elements = new ObjectGroupDataElementData.Builder().Build(rootNode);
            objectDataExGuidList.AddRange(
                        elements.Where(element => element.DataElementType == DataElementType.ObjectGroupDataElementData)
                                .Select(element => element.DataElementExtendedGUID)
                                .ToArray());
           
            return elements;
        }

        /// <summary>
        /// This method is used to create the revision manifest data element.
        /// </summary>
        /// <param name="rootObjectExGuid">Specify the root node object extended GUID.</param>
        /// <param name="baseRevisionID">Specify the base revision Id.</param>
        /// <param name="refferenceObjectDataExGuidList">Specify the reference object data extended list.</param>
        /// <param name="revisionMapping">Input/output parameter to represent the mapping of revision manifest.</param>
        /// <param name="currentRevisionID">Output parameter to represent the revision GUID.</param>
        /// <returns>Return the revision manifest data element.</returns>
        public static DataElement CreateRevisionManifestDataElement(ExGuid rootObjectExGuid, ExGuid baseRevisionID, List<ExGuid> refferenceObjectDataExGuidList, ref Dictionary<ExGuid, ExGuid> revisionMapping, out ExGuid currentRevisionID)
        {
            RevisionManifestDataElementData data = new RevisionManifestDataElementData();
            data.RevisionManifest.RevisionID = new ExGuid(1u, Guid.NewGuid());
            data.RevisionManifest.BaseRevisionID = new ExGuid(baseRevisionID);

            // Set the root object data ExGuid
            data.RevisionManifestRootDeclareList.Add(new RevisionManifestRootDeclare() { RootExtendedGUID = new ExGuid(2u, RootExGuid), ObjectExtendedGUID = new ExGuid(rootObjectExGuid) });

            // Set all the reference object data
            if (refferenceObjectDataExGuidList != null)
            {
                foreach (ExGuid dataGuid in refferenceObjectDataExGuidList)
                {
                    data.RevisionManifestObjectGroupReferencesList.Add(new RevisionManifestObjectGroupReferences(dataGuid));
                }
            }

            DataElement dataElement = new DataElement(DataElementType.RevisionManifestDataElementData, data);
            revisionMapping.Add(data.RevisionManifest.RevisionID, dataElement.DataElementExtendedGUID);
            currentRevisionID = data.RevisionManifest.RevisionID;
            return dataElement;
        }

        /// <summary>
        /// This method is used to create the cell manifest data element.
        /// </summary>
        /// <param name="revisionId">Specify the revision GUID.</param>
        /// <param name="cellIDMapping">Input/output parameter to represent the mapping of cell manifest.</param>
        /// <returns>Return the cell manifest data element.</returns>
        public static DataElement CreateCellMainifestDataElement(ExGuid revisionId, ref Dictionary<CellID, ExGuid> cellIDMapping)
        {
            CellManifestDataElementData data = new CellManifestDataElementData();
            data.CellManifestCurrentRevision = new CellManifestCurrentRevision() { CellManifestCurrentRevisionExtendedGUID = new ExGuid(revisionId) };
            DataElement dataElement = new DataElement(DataElementType.CellManifestDataElementData, data);

            CellID cellID = new CellID(new ExGuid(1u, RootExGuid), new ExGuid(1u, CellSecondExGuid));
            cellIDMapping.Add(cellID, dataElement.DataElementExtendedGUID);
            return dataElement;
        }

        /// <summary>
        /// This method is used to create the storage manifest data element.
        /// </summary>
        /// <param name="cellIDMapping">Specify the mapping of cell manifest.</param>
        /// <returns>Return the storage manifest data element.</returns>
        public static DataElement CreateStorageManifestDataElement(Dictionary<CellID, ExGuid> cellIDMapping)
        {
            StorageManifestDataElementData data = new StorageManifestDataElementData();
            data.StorageManifestSchemaGUID = new StorageManifestSchemaGUID() { GUID = SchemaGuid };

            foreach (KeyValuePair<CellID, ExGuid> kv in cellIDMapping)
            {
                StorageManifestRootDeclare manifestRootDeclare = new StorageManifestRootDeclare();
                manifestRootDeclare.RootExtendedGUID = new ExGuid(2u, RootExGuid);
                manifestRootDeclare.CellID = new CellID(kv.Key);
                data.StorageManifestRootDeclareList.Add(manifestRootDeclare);
            }

            return new DataElement(DataElementType.StorageManifestDataElementData, data);
        }

        /// <summary>
        /// This method is used to create the storage index data element.
        /// </summary>
        /// <param name="manifestExGuid">Specify the storage manifest data element extended GUID.</param>
        /// <param name="cellIDMappings">Specify the mapping of cell manifest.</param>
        /// <param name="revisionIDMappings">Specify the mapping of revision manifest.</param>
        /// <returns>Return the storage index data element.</returns>
        public static DataElement CreateStorageIndexDataElement(ExGuid manifestExGuid, Dictionary<CellID, ExGuid> cellIDMappings, Dictionary<ExGuid, ExGuid> revisionIDMappings)
        {
            StorageIndexDataElementData data = new StorageIndexDataElementData();

            data.StorageIndexManifestMapping = new StorageIndexManifestMapping();
            data.StorageIndexManifestMapping.ManifestMappingExtendedGUID = new ExGuid(manifestExGuid);
            data.StorageIndexManifestMapping.ManifestMappingSerialNumber = new SerialNumber(System.Guid.NewGuid(), SequenceNumberGenerator.GetCurrentSerialNumber());

            foreach (KeyValuePair<CellID, ExGuid> kv in cellIDMappings)
            {
                StorageIndexCellMapping cellMapping = new StorageIndexCellMapping();
                cellMapping.CellID = kv.Key;
                cellMapping.CellMappingExtendedGUID = kv.Value;
                cellMapping.CellMappingSerialNumber = new SerialNumber(System.Guid.NewGuid(), SequenceNumberGenerator.GetCurrentSerialNumber());
                data.StorageIndexCellMappingList.Add(cellMapping);
            }

            foreach (KeyValuePair<ExGuid, ExGuid> kv in revisionIDMappings)
            {
                StorageIndexRevisionMapping revisionMapping = new StorageIndexRevisionMapping();
                revisionMapping.RevisionExtendedGUID = kv.Key;
                revisionMapping.RevisionMappingExtendedGUID = kv.Value;
                revisionMapping.RevisionMappingSerialNumber = new SerialNumber(Guid.NewGuid(), SequenceNumberGenerator.GetCurrentSerialNumber());
                data.StorageIndexRevisionMappingList.Add(revisionMapping);
            }

            return new DataElement(DataElementType.StorageIndexDataElementData, data);
        }

        /// <summary>
        /// This method is used to get the list of object group data element from a list of data element.
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="storageIndexExGuid">Specify the storage index extended GUID.</param>
        /// <param name="rootExGuid">Output parameter to represent the root node object.</param>
        /// <returns>Return the list of object group data elements.</returns>
        public static List<ObjectGroupDataElementData> GetDataObjectDataElementData(List<DataElement> dataElements, ExGuid storageIndexExGuid, out ExGuid rootExGuid)
        {
            ExGuid manifestMappingGuid;
            Dictionary<CellID, ExGuid> cellIDMappings;
            Dictionary<ExGuid, ExGuid> revisionIDMappings;
            AnalyzeStorageIndexDataElement(dataElements, storageIndexExGuid, out manifestMappingGuid, out cellIDMappings, out revisionIDMappings);
            StorageManifestDataElementData manifestData = GetStorageManifestDataElementData(dataElements, manifestMappingGuid);
            if (manifestData == null)
            {
                throw new InvalidOperationException("Cannot find the storage manifest data element with ExGuid " + manifestMappingGuid.GUID.ToString());
            }
            
            CellManifestDataElementData cellData = GetCellManifestDataElementData(dataElements, manifestData, cellIDMappings);
            RevisionManifestDataElementData revisionData = GetRevisionManifestDataElementData(dataElements, cellData, revisionIDMappings);
            return GetDataObjectDataElementData(dataElements, revisionData, out rootExGuid);
        }

        /// <summary>
        /// This method is used to try to analyze the returned whether data elements are complete.
        /// </summary>
        /// <param name="dataElements">Specify the data elements list.</param>
        /// <param name="storageIndexExGuid">Specify the storage index extended GUID.</param>
        /// <returns>If the data elements start with the specified storage index extended GUID are complete, return true. Otherwise return false.</returns>
        public static bool TryAnalyzeWhetherFullDataElementList(List<DataElement> dataElements, ExGuid storageIndexExGuid)
        {
            ExGuid manifestMappingGuid;
            Dictionary<CellID, ExGuid> cellIDMappings;
            Dictionary<ExGuid, ExGuid> revisionIDMappings;
            if (!AnalyzeStorageIndexDataElement(dataElements, storageIndexExGuid, out manifestMappingGuid, out cellIDMappings, out revisionIDMappings))
            {
                return false;
            }

            if (cellIDMappings.Count == 0)
            {
                return false;
            }

            if (revisionIDMappings.Count == 0)
            {
                return false;
            }

            StorageManifestDataElementData manifestData = GetStorageManifestDataElementData(dataElements, manifestMappingGuid);
            if (manifestData == null)
            {
                return false;
            }

            foreach (StorageManifestRootDeclare kv in manifestData.StorageManifestRootDeclareList)
            {
                if (!cellIDMappings.ContainsKey(kv.CellID))
                {
                    throw new InvalidOperationException(string.Format("Cannot fin the Cell ID {0} in the cell id mapping", kv.CellID.ToString()));
                }

                ExGuid cellMappingID = cellIDMappings[kv.CellID];
                DataElement dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(cellMappingID));
                if (dataElement == null)
                {
                    return false;
                }

                CellManifestDataElementData cellData = dataElement.GetData<CellManifestDataElementData>();
                ExGuid currentRevisionExGuid = cellData.CellManifestCurrentRevision.CellManifestCurrentRevisionExtendedGUID;
                if (!revisionIDMappings.ContainsKey(currentRevisionExGuid))
                {
                    throw new InvalidOperationException(string.Format("Cannot find the revision id {0} in the revisionMapping", currentRevisionExGuid.ToString()));
                }

                ExGuid revisionMapping = revisionIDMappings[currentRevisionExGuid];
                dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(revisionMapping));
                if (dataElement == null)
                {
                    return false;
                }

                RevisionManifestDataElementData revisionData = dataElement.GetData<RevisionManifestDataElementData>();
                foreach (RevisionManifestObjectGroupReferences reference in revisionData.RevisionManifestObjectGroupReferencesList)
                {
                    dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(reference.ObjectGroupExtendedGUID));
                    if (dataElement == null)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// This method is used to analyze whether the data elements are confirmed to the schema defined in MS-FSSHTTPD. 
        /// </summary>
        /// <param name="dataElements">Specify the data elements list.</param>
        /// <param name="storageIndexExGuid">Specify the storage index extended GUID.</param>
        /// <returns>If the data elements confirms to the schema defined in the MS-FSSHTTPD returns true, otherwise false.</returns>
        public static bool TryAnalyzeWhetherConfirmSchema(List<DataElement> dataElements, ExGuid storageIndexExGuid)
        {
            DataElement storageIndexDataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(storageIndexExGuid));
            if (storageIndexExGuid == null)
            {
                return false;
            }

            StorageIndexDataElementData storageIndexData = storageIndexDataElement.GetData<StorageIndexDataElementData>();
            ExGuid manifestMappingGuid = storageIndexData.StorageIndexManifestMapping.ManifestMappingExtendedGUID;

            DataElement storageManifestDataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(manifestMappingGuid));
            if (storageManifestDataElement == null)
            {
                return false;
            }

            return SchemaGuid.Equals(storageManifestDataElement.GetData<StorageManifestDataElementData>().StorageManifestSchemaGUID.GUID);
        }

        /// <summary>
        /// This method is used to analyze the storage index data element to get all the mappings. 
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="storageIndexExGuid">Specify the storage index extended GUID.</param>
        /// <param name="manifestMappingGuid">Output parameter to represent the storage manifest mapping GUID.</param>
        /// <param name="cellIDMappings">Output parameter to represent the mapping of cell id.</param>
        /// <param name="revisionIDMappings">Output parameter to represent the revision id.</param>
        /// <returns>Return true if analyze the storage index succeeds, otherwise return false.</returns>
        public static bool AnalyzeStorageIndexDataElement(
                        List<DataElement> dataElements, 
                        ExGuid storageIndexExGuid, 
                        out ExGuid manifestMappingGuid,
                        out Dictionary<CellID, ExGuid> cellIDMappings, 
                        out Dictionary<ExGuid, ExGuid> revisionIDMappings)
        {
            manifestMappingGuid = null;
            cellIDMappings = null;
            revisionIDMappings = null;

            if (storageIndexExGuid == null)
            {
                return false;
            }

            DataElement storageIndexDataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(storageIndexExGuid));
            StorageIndexDataElementData storageIndexData = storageIndexDataElement.GetData<StorageIndexDataElementData>();
            manifestMappingGuid = storageIndexData.StorageIndexManifestMapping.ManifestMappingExtendedGUID;

            cellIDMappings = new Dictionary<CellID, ExGuid>();
            foreach (StorageIndexCellMapping kv in storageIndexData.StorageIndexCellMappingList)
            {
                cellIDMappings.Add(kv.CellID, kv.CellMappingExtendedGUID);
            }

            revisionIDMappings = new Dictionary<ExGuid, ExGuid>();
            foreach (StorageIndexRevisionMapping kv in storageIndexData.StorageIndexRevisionMappingList)
            {
                revisionIDMappings.Add(kv.RevisionExtendedGUID, kv.RevisionMappingExtendedGUID);
            }

            return true;
        }

        /// <summary>
        /// This method is used to get storage manifest data element from a list of data element.
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="manifestMapping">Specify the manifest mapping GUID.</param>
        /// <returns>Return the storage manifest data element.</returns>
        public static StorageManifestDataElementData GetStorageManifestDataElementData(List<DataElement> dataElements, ExGuid manifestMapping)
        {
            DataElement storageManifestDataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(manifestMapping));
            if (storageManifestDataElement == null)
            {
                return null;
            }

            return storageManifestDataElement.GetData<StorageManifestDataElementData>();    
        }

        /// <summary>
        /// This method is used to get cell manifest data element from a list of data element.
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="manifestDataElementData">Specify the manifest data element.</param>
        /// <param name="cellIDMappings">Specify mapping of cell id.</param>
        /// <returns>Return the cell manifest data element.</returns>
        public static CellManifestDataElementData GetCellManifestDataElementData(List<DataElement> dataElements, StorageManifestDataElementData manifestDataElementData, Dictionary<CellID, ExGuid> cellIDMappings)
        {
            CellID cellID = new CellID(new ExGuid(1u, RootExGuid), new ExGuid(1u, CellSecondExGuid));

            foreach (StorageManifestRootDeclare kv in manifestDataElementData.StorageManifestRootDeclareList)
            {
                if (kv.RootExtendedGUID.Equals(new ExGuid(2u, RootExGuid)) && kv.CellID.Equals(cellID))
                {
                    if (!cellIDMappings.ContainsKey(kv.CellID))
                    {
                        throw new InvalidOperationException(string.Format("Cannot fin the Cell ID {0} in the cell id mapping", cellID.ToString()));
                    }

                    ExGuid cellMappingID = cellIDMappings[kv.CellID];

                    DataElement dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(cellMappingID));
                    if (dataElement == null)
                    {
                        throw new InvalidOperationException("Cannot find the  cell data element with ExGuid " + cellMappingID.GUID.ToString());
                    }

                    return dataElement.GetData<CellManifestDataElementData>();
                }
            }

            throw new InvalidOperationException("Cannot find the CellManifestDataElement");
        }

        /// <summary>
        /// This method is used to get revision manifest data element from a list of data element.
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="cellData">Specify the cell data element.</param>
        /// <param name="revisionIDMappings">Specify mapping of revision id.</param>
        /// <returns>Return the revision manifest data element.</returns>
        public static RevisionManifestDataElementData GetRevisionManifestDataElementData(List<DataElement> dataElements, CellManifestDataElementData cellData, Dictionary<ExGuid, ExGuid> revisionIDMappings)
        {
            ExGuid currentRevisionExGuid = cellData.CellManifestCurrentRevision.CellManifestCurrentRevisionExtendedGUID;

            if (!revisionIDMappings.ContainsKey(currentRevisionExGuid))
            {
                throw new InvalidOperationException(string.Format("Cannot find the revision id {0} in the revisionMapping", currentRevisionExGuid.ToString()));
            }
            
            ExGuid revisionMapping = revisionIDMappings[currentRevisionExGuid];

            DataElement dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(revisionMapping));
            if (dataElement == null)
            {
                throw new InvalidOperationException("Cannot find the revision data element with ExGuid " + revisionMapping.GUID.ToString());
            }

            return dataElement.GetData<RevisionManifestDataElementData>();
        }

        /// <summary>
        /// This method is used to get a list of object group data element from a list of data element.
        /// </summary>
        /// <param name="dataElements">Specify the data element list.</param>
        /// <param name="revisionData">Specify the revision data.</param>
        /// <param name="rootExGuid">Specify the root node object extended GUID.</param>
        /// <returns>Return the list of object group data element.</returns>
        public static List<ObjectGroupDataElementData> GetDataObjectDataElementData(List<DataElement> dataElements, RevisionManifestDataElementData revisionData, out ExGuid rootExGuid)
        {
            rootExGuid = null;

            foreach (RevisionManifestRootDeclare kv in revisionData.RevisionManifestRootDeclareList)
            {
                if (kv.RootExtendedGUID.Equals(new ExGuid(2u, RootExGuid)))
                {
                    rootExGuid = kv.ObjectExtendedGUID;
                    break;
                }
            }

            List<ObjectGroupDataElementData> dataList = new List<ObjectGroupDataElementData>();

            foreach (RevisionManifestObjectGroupReferences kv in revisionData.RevisionManifestObjectGroupReferencesList)
            {
                DataElement dataElement = dataElements.Find(element => element.DataElementExtendedGUID.Equals(kv.ObjectGroupExtendedGUID));
                if (dataElement == null)
                {
                    throw new InvalidOperationException("Cannot find the object group data element with ExGuid " + kv.ObjectGroupExtendedGUID.GUID.ToString());
                }

                dataList.Add(dataElement.GetData<ObjectGroupDataElementData>());
            }

            return dataList;
        }
    }
}