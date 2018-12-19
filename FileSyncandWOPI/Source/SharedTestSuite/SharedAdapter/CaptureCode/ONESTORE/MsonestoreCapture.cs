namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using Microsoft.Protocols.TestTools;
    using System;

    public class MsonestoreCapture
    {
        public void Validate(MSOneStorePackage instance, ITestSite site)
        {
            this.VerifyStorageManifest(instance.StorageManifest, site);
            foreach (RevisionManifestDataElementData revisionManifest in instance.RevisionManifests)
            {
                this.VerifyRevisions(revisionManifest, site);
            }

            foreach(RevisionStoreObject revisionStoreObject in instance.RevisionStoreObjects)
            {
                this.VerifyRevisionStoreObject(revisionStoreObject, site);
            }
        }

        private void VerifyStorageManifest(StorageManifestDataElementData instance, ITestSite site)
        {

        }

        private void VerifyRevisions(RevisionManifestDataElementData instance, ITestSite site)
        {

        }

        private void VerifyRevisionStoreObject(RevisionStoreObject instance, ITestSite site)
        {
            
        }
    }
}
