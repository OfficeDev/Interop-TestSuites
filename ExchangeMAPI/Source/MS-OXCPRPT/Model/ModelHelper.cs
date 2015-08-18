namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using Microsoft.Modeling;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Helper class of Model
    /// </summary>
    public static class ModelHelper
    {
        /// <summary>
        /// Requirement capture
        /// </summary>
        /// <param name="id">Requirement id</param>
        /// <param name="description">Requirement description</param>
        public static void CaptureRequirement(int id, string description)
        {
            Requirement.Capture(RequirementId.Make("MS-OXCPRPT", id, description));
        }
    }
}