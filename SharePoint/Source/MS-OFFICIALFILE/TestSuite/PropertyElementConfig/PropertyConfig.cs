namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using System;
    using System.Collections.Generic;
    using System.Xml.Serialization;

    #region Properties configuration class

    /// <summary>
    /// The class define the properties configure class which is corresponding with the definition in the PropertyConfig.xml.
    /// </summary>
    [Serializable]
    [XmlType(Namespace = "http://schemas.microsoft.com/sharepoint/soap/recordsrepository/")]
    public class PropertyConfig
    {
        /// <summary>
        /// The records repository properties.
        /// </summary>
        private List<RecordsRepositoryProperty> recordsRepositoryProperties;

        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyConfig" /> class.
        /// </summary>
        public PropertyConfig()
        {
            this.recordsRepositoryProperties = new List<RecordsRepositoryProperty>();
        }

        /// <summary>
        /// Gets or sets the records repository properties. 
        /// </summary>
        public List<RecordsRepositoryProperty> RecordsRepositoryProperties
        {
            get { return this.recordsRepositoryProperties; }
            set { this.recordsRepositoryProperties = value; }
        }
    }
    #endregion 
}