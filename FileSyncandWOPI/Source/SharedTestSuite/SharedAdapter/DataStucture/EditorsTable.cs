namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;

    /// <summary>
    /// The class is used to represent the editors table.
    /// </summary>
    [Serializable]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class EditorsTable
    {
        /// <summary>
        /// Gets or sets an array of editors. 
        /// </summary>
        public Editor[] Editors { get; set; }
    }
}