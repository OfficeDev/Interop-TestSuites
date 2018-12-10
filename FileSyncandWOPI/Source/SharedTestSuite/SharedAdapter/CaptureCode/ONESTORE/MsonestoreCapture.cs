namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestTools;
    using System;

    public class MsonestoreCapture
    {
        public void Validate(DataElement instance, ITestSite site)
        {
            if (instance.Data is StorageManifestDataElementData)
            {
                this.VerifyStorageManifest((StorageManifestDataElementData)instance.Data, site);
            }
            else if (instance.Data is RevisionManifestDataElementData)
            {
                this.VerifyRevisions((RevisionManifestDataElementData)instance.Data, site);
            }
            else if (instance.Data is ObjectGroupDataElementData)
            {
                this.VerifyObjects((ObjectGroupDataElementData)instance.Data, site);
            }
        }

        private void VerifyStorageManifest(StorageManifestDataElementData instance, ITestSite site)
        {
            site.Assert.AreEqual<Guid>(Guid.Parse("{1F937CB4-B26F-445F-B9F8-17E20160E461}"), instance.StorageManifestSchemaGUID.GUID, "The GUID should be {1F937CB4-B26F-445F-B9F8-17E20160E461}");
        }

        private void VerifyRevisions(RevisionManifestDataElementData instance, ITestSite site)
        {

        }

        private void VerifyObjects(ObjectGroupDataElementData instance, ITestSite site)
        {
            if (instance.ObjectGroupDeclarations.ObjectDeclarationList.Count > 0
                && instance.ObjectGroupData.ObjectGroupObjectDataList.Count > 0)
            {
                int count = instance.ObjectGroupDeclarations.ObjectDeclarationList.Count;

                for (int i = 0; i < count; i++)
                {
                    ObjectGroupObjectDeclare objectDeclaration = instance.ObjectGroupDeclarations.ObjectDeclarationList[i];
                    ObjectGroupObjectData objectData = instance.ObjectGroupData.ObjectGroupObjectDataList[i];
                    this.VerifyObjectData(objectDeclaration, objectData, site);
                }
            }
            else if (instance.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList.Count > 0
                && instance.ObjectGroupData.ObjectGroupObjectDataBLOBReferenceList.Count > 0)
            {
                int count = instance.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList.Count;

                for (int i = 0; i < count; i++)
                {
                    ObjectGroupObjectBLOBDataDeclaration objectGroupObjectBLOBDataDeclaration = instance.ObjectGroupDeclarations.ObjectGroupObjectBLOBDataDeclarationList[i];
                    ObjectGroupObjectDataBLOBReference objectGroupObjectDataBLOBReference = instance.ObjectGroupData.ObjectGroupObjectDataBLOBReferenceList[i];
                    this.VerifyObjectDataBLOB(objectGroupObjectBLOBDataDeclaration, objectGroupObjectDataBLOBReference, site);
                }
            }
        }

        private void VerifyObjectData(ObjectGroupObjectDeclare objectDeclaration, ObjectGroupObjectData objectData, ITestSite sit)
        {

        }

        private void VerifyObjectDataBLOB(ObjectGroupObjectBLOBDataDeclaration objectDataBLOBDeclaration, ObjectGroupObjectDataBLOBReference objectDataBLOBReference, ITestSite sit)
        {

        }
    }
}
