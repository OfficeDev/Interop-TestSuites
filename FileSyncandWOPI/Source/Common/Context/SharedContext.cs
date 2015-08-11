namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class represents context which will be used to shared necessary information for both the MS-FSSHTTP and MS-WOPI.
    /// </summary>
    public class SharedContext
    {
        /// <summary>
        /// Specify the thread local storage context.
        /// </summary>
        [ThreadStatic]
        private static SharedContext current;

        /// <summary>
        /// Specify the properties stored in the context.
        /// </summary>
        private Dictionary<string, object> properties;

        /// <summary>
        /// Prevents a default instance of the SharedContext class from being created.
        /// </summary>
        private SharedContext()
        {
            this.properties = new Dictionary<string, object>();
        }

        /// <summary>
        /// Gets the current context stored in the thread local storage.
        /// </summary>
        public static SharedContext Current
        {
            get
            {
                if (current == null)
                {
                    current = new SharedContext();
                }

                return current;
            }
        }

        /// <summary>
        /// Gets or sets an object provides logging, assertions, and SUT adapters for test code onto its execution context.
        /// </summary>
        public ITestSite Site
        {
            get
            {
                return this.GetValueOrDefault<ITestSite>("Site");
            }

            set
            {
                this.AddOrUpdate("Site", value);
            }
        }

        /// <summary>
        /// Gets or sets the operation type which will indicate whether the operations are defined in MS-WOPI and MS-FSSHTTP.
        /// </summary>
        public OperationType OperationType
        {
            get
            {
                return this.GetValueOrDefault<OperationType>("OperationType", OperationType.FSSHTTPCellStorageRequest);
            }

            set
            {
                this.AddOrUpdate("OperationType", value);
            }
        }

        /// <summary>
        /// Gets or sets the VersionType defined in the MS-FSSHTTP.
        /// </summary>
        public VersionType CellStorageVersionType
        {
            get
            {
                return this.GetValueOrDefault<VersionType>("CellStorageVersionType");
            }

            set
            {
                this.AddOrUpdate("CellStorageVersionType", value);
            }
        }

        /// <summary>
        /// Gets or sets request target Url.
        /// </summary>
        public string TargetUrl
        {
            get
            {
                return this.GetValueOrDefault<string>("TargetUrl");
            }

            set
            {
                this.AddOrUpdate("TargetUrl", value);
            }
        }

        /// <summary>
        /// Gets or sets the endpoint configuration name.
        /// </summary>
        public string EndpointConfigurationName
        {
            get
            {
                return this.GetValueOrDefault<string>("EndpointConfigurationName");
            }

            set
            {
                this.AddOrUpdate("EndpointConfigurationName", value);
            }
        }

        /// <summary>
        /// Gets or sets the user name.
        /// </summary>
        public string UserName
        {
            get
            {
                return this.GetValueOrDefault<string>("UserName");
            }

            set
            {
                this.AddOrUpdate("UserName", value);
            }
        }

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
        public string Password
        {
            get
            {
                return this.GetValueOrDefault<string>("Password");
            }

            set
            {
                this.AddOrUpdate("Password", value);
            }
        }

        /// <summary>
        /// Gets or sets the domain.
        /// </summary>
        public string Domain
        {
            get
            {
                return this.GetValueOrDefault<string>("Domain");
            }

            set
            {
                this.AddOrUpdate("Domain", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-ProofOld header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPIProofOld 
        {
            get
            {
                return this.GetValueOrDefault<string>("XWOPIProofOld");
            }

            set
            {
                this.AddOrUpdate("XWOPIProofOld", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-Proof header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPIProof
        {
            get
            {
                return this.GetValueOrDefault<string>("XWOPIProof");
            }

            set
            {
                this.AddOrUpdate("XWOPIProof", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-TimeStamp header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPITimeStamp 
        {
            get
            {
                return this.GetValueOrDefault<string>("XWOPITimeStamp");
            }

            set
            {
                this.AddOrUpdate("XWOPITimeStamp", value);
            }
        }

        /// <summary>
        /// Gets or sets the Authorization header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPIAuthorization
        {
            get
            {
                return this.GetValueOrDefault<string>("XWOPIAuthorization");
            }

            set
            {
                this.AddOrUpdate("XWOPIAuthorization", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-RelativeTarget header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPIRelativeTarget
        {
            get
            {
                return this.GetValueOrDefault<string>("X-WOPI-RelativeTarget");
            }

            set
            {
                this.AddOrUpdate("X-WOPI-RelativeTarget", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-Override header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPIOverride
        {
            get
            {
                return this.GetValueOrDefault<string>("X-WOPI-Override");
            }

            set
            {
                this.AddOrUpdate("X-WOPI-Override", value);
            }
        }

        /// <summary>
        /// Gets or sets the X-WOPI-Size header value defined in the MS-WOPI.
        /// </summary>
        public string XWOPISize
        {
            get
            {
                return this.GetValueOrDefault<string>("X-WOPI-Size");
            }

            set
            {
                this.AddOrUpdate("X-WOPI-Size", value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the X-WOPI-RelativeTarget header is send. If true, the channel will send optional X-WOPI-RelativeTarget header. Otherwise it will not send.
        /// The default value is true if it is not set.
        /// </summary>
        public bool IsXWOPIRelativeTargetSpecified
        {
            get
            {
                return this.GetValueOrDefault<bool>("IsXWOPIRelativeTargetSpecified", true);
            }

            set
            {
                this.AddOrUpdate("IsXWOPIRelativeTargetSpecified", value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the X-WOPI-Override header is send. If true, the channel will send optional X-WOPI-Override header. Otherwise it will not send.
        /// The default value is true if it is not set.
        /// </summary>
        public bool IsXWOPIOverrideSpecified
        {
            get
            {
                return this.GetValueOrDefault<bool>("IsXWOPIOverrideSpecified", true);
            }

            set
            {
                this.AddOrUpdate("IsXWOPIOverrideSpecified", value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the X-WOPI-Size header is send. If true, the channel will send optional X-WOPI-Size header. Otherwise it will not send.
        /// The default value is true if it is not set.
        /// </summary>
        public bool IsXWOPISizeSpecified
        {
            get
            {
                return this.GetValueOrDefault<bool>("IsXWOPISizeSpecified", true);
            }

            set
            {
                this.AddOrUpdate("IsXWOPISizeSpecified", value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether MS-FSSHTTP related requirements will be captured. If the OperationType is FSSHTTPCellStorageRequest <see cref="TestSuites.Common.OperationType"/>, it will always return true.
        /// Otherwise, the default value is false if it is not set.
        /// </summary>
        public bool IsMsFsshttpRequirementsCaptured
        {
            get
            {
                if (this.OperationType == TestSuites.Common.OperationType.FSSHTTPCellStorageRequest)
                {
                    return true;
                }

                return this.GetValueOrDefault<bool>("IsMsFsshttpRequirementsCaptured", false);
            }

            set
            {
                this.AddOrUpdate("IsMsFsshttpRequirementsCaptured", value);
            }
        }

        /// <summary>
        /// This method is used to get the value with the specified key if it exists.
        /// If the key does not exist, the default(T) value will be returned.
        /// </summary>
        /// <typeparam name="T">Specify the type of the value.</typeparam>
        /// <param name="key">Specify the key.</param>
        /// <returns>Return the value associated with the key.</returns>
        public T GetValueOrDefault<T>(string key)
        {
            return this.GetValueOrDefault<T>(key, default(T));
        }

        /// <summary>
        /// This method is used to get the value with the specified key if it exists.
        /// If the key does not exist, the defaultValue value will be returned. 
        /// </summary>
        /// <typeparam name="T">Specify the type of the value.</typeparam>
        /// <param name="key">Specify the key.</param>
        /// <param name="defaultValue">Specify the default value.</param>
        /// <returns>Return the value associated with the key.</returns>
        public T GetValueOrDefault<T>(string key, T defaultValue)
        {
            object outValue;
            if (!this.properties.TryGetValue(key, out outValue))
            {
                return defaultValue;
            }

            return (T)outValue;
        }

        /// <summary>
        /// This method is used to create or update the entry with the specified key.
        /// </summary>
        /// <param name="key">Specify the key.</param>
        /// <param name="value">Specify the value.</param>
        public void AddOrUpdate(string key, object value)
        {
            if (!this.properties.ContainsKey(key))
            {
                this.properties.Add(key, value);
            }
            else
            {
                this.properties[key] = value;
            }
        }

        /// <summary>
        /// Clear all the current context properties.
        /// </summary>
        public void Clear()
        {
            this.properties.Clear();
        }
    }
}