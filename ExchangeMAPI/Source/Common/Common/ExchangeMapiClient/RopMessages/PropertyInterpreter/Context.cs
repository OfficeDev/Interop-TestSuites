namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Collections.Generic;

    /// <summary>
    /// Context contains all needed information to interpret properties:
    /// propertyBytes: Input bytes, properties: columns of properties to parse, 
    /// curIndex: Current propertyBytes' index
    /// curPropertyRow: Current property row, curProperty: current property to parse,
    /// propertyRows: Parsing result consisted by PropertyRows
    /// </summary>
    public class Context
    {
        /// <summary>
        /// Context instance.
        /// </summary>
        private static Context instance;

        /// <summary>
        /// Bytes to parse
        /// </summary>
        private byte[] propertyBytes;

        /// <summary>
        /// Columns of properties to parse
        /// </summary>
        private List<Property> properties;

        /// <summary>
        /// Current parsing index of propertyBytes
        /// </summary>
        private int curIndex;

        /// <summary>
        /// Current parsing property
        /// </summary>
        private Property curProperty;

        /// <summary>
        /// Current parsing property row
        /// </summary>
        private PropertyRow curPropertyRow;

        /// <summary>
        /// PropertyRow that already parsed
        /// </summary>
        private List<PropertyRow> propertyRows;

        /// <summary>
        /// Prevents a default instance of the <see cref="Context"/> class from being created.
        /// </summary>
        private Context()
        {
            this.propertyRows = new List<PropertyRow>();
        }

        /// <summary>
        /// Gets context instance.
        /// </summary>
        public static Context Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Context();
                }

                return instance;
            }
        }

        #region Properties
        /// <summary>
        /// Gets or sets the bytes of the properties.
        /// </summary>
        public byte[] PropertyBytes
        {
            get
            {
                return this.propertyBytes;
            }

            set
            {
                this.propertyBytes = value;
            }
        }

        /// <summary>
        /// Gets or sets properties.
        /// </summary>
        public List<Property> Properties
        {
            get
            {
                return this.properties;
            }

            set
            {
                this.properties = value;
            }
        }

        /// <summary>
        /// Gets or sets current index of the property bytes parsed.
        /// </summary>
        public int CurIndex
        {
            get
            {
                return this.curIndex;
            }

            set
            {
                this.curIndex = value;
            }
        }

        /// <summary>
        /// Gets below properties is used by PropertyInterpreter internally 
        /// </summary>
        public PropertyRow CurPropertyRow
        {
            get
            {
                return this.curPropertyRow;
            }
        }

        /// <summary>
        /// Gets or sets current property parsed.
        /// </summary>
        public Property CurProperty
        {
            get
            {
                return this.curProperty;
            }

            set
            {
                this.curProperty = value;
            }
        }

        /// <summary>
        /// Gets property rows.
        /// </summary>
        public List<PropertyRow> PropertyRows
        {
            get
            {
                return this.propertyRows;
            }
        }

        #endregion

        /// <summary>
        /// Initialize the instance.
        /// </summary>
        public void Init()
        {
            this.propertyBytes = null;

            this.curProperty = null;
            this.curPropertyRow = null;
            this.curIndex = 0;

            this.properties = new List<Property>();
            this.propertyRows = new List<PropertyRow>();
        }

        /// <summary>
        /// Indicates input bytes is processed to the end
        /// </summary>
        /// <returns>Return true indicates input bytes already reach ending, otherwise false </returns>
        public bool IsEnd()
        {
            return this.curIndex >= this.propertyBytes.Length;
        }

        /// <summary>
        /// Return the count of available bytes of input bytes
        /// </summary>
        /// <returns>The count in bytes of not processed input bytes</returns>
        public int AvailBytes()
        {
            return this.propertyBytes.Length - this.curIndex;
        }
    }
}