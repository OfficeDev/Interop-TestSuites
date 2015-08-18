namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Specifies that one class is Serializable.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class |
            AttributeTargets.Enum |
            AttributeTargets.Struct)]
    public sealed class SerializableObjectAttribute : Attribute
    {
        /// <summary>
        /// Use SelfSerialize.
        /// </summary>
        private bool useSelfSerialize;

        /// <summary>
        /// Use SelfDeserialize
        /// </summary>
        private bool useSelfDeserialize;

        /// <summary>
        /// Initializes a new instance of the SerializableObjectAttribute class.
        /// </summary>
        /// <param name="useSelfSerialize">Whether to use self-Serialize method.</param>
        /// <param name="useSelfDeserialize">Whether to use self deserialize method.</param>
        public SerializableObjectAttribute(bool useSelfSerialize, bool useSelfDeserialize)
        {
            this.useSelfSerialize = useSelfSerialize;
            this.useSelfDeserialize = useSelfDeserialize;
        }

        /// <summary>
        /// Gets or sets a value indicating whether to use self-Serialize method.
        /// </summary>
        public bool UseSelfSerialize
        {
            get { return this.useSelfSerialize; }
            set { this.useSelfSerialize = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to use self-deserialize method.
        /// </summary>
        public bool UseSelfDeserialize
        {
            get { return this.useSelfDeserialize; }
            set { this.useSelfDeserialize = value; }
        }

        /// <summary>
        /// Indicate whether a class has a SerializableObject attribute.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <returns>If the object has a SerializableObject attribute,
        /// return true, else false.
        /// </returns>
        public static bool IsSerializableObject(object obj)
        {
            if (obj != null)
            {
                Type type = obj.GetType();
                object[] attrs = type.GetCustomAttributes(
                    typeof(SerializableObjectAttribute),
                    false);
                return attrs.Length > 0;
            }

            return false;
        }

        /// <summary>
        /// Reflect the first SerializableObject from an object.
        /// </summary>
        /// <param name="obj">An object instance.</param>
        /// <returns>The first SerializableObject attribute of the object.</returns>
        public static SerializableObjectAttribute GetSerializableObject(object obj)
        {
            if (!(obj is SerializableBase))
            {
                AdapterHelper.Site.Assert.Fail("The object 'obj' is not a SerializableBase instance.");
            }

            Type type = obj.GetType();
            object[] attrs = type.GetCustomAttributes(
                typeof(SerializableObjectAttribute),
                false);
            return attrs[0] as SerializableObjectAttribute;
        }
    }
}