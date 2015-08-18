namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Helper class of Model
    /// </summary>
    public static class ModelHelper
    {
        #region Requirement Capture
        /// <summary>
        /// Requirement Capture in model 
        /// </summary>
        /// <param name="id">Requirement ID</param>
        /// <param name="description">Requirement Description</param>
        public static void CaptureRequirement(int id, string description)
        {
            Requirement.Capture(RequirementId.Make("MS-OXCTABL", id, description));
        }
        #endregion
    }
}