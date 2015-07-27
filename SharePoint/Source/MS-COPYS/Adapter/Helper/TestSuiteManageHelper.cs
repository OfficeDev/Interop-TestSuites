//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;
    using System.Net;
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class used to support required functions for test class initialization and clean up.
    /// </summary>
    public class TestSuiteManageHelper
    {
        #region fields
        /// <summary>
        /// A string represents the password for the default user.
        /// </summary>
        private static string passwordOfDefaultUser;

        /// <summary>
        /// A string represents the domain name for the default user.
        /// </summary>
        private static string domainOfDefaultUser;

        /// <summary>
        /// A string represents the user name for the default user.
        /// </summary>
        private static string defaultUser;

        /// <summary>
        ///  A SoapVersion enum represents the current SOAP version used by CopySoap web service.
        /// </summary>
        private static SoapVersion currentSoapVersion;

        /// <summary>
        /// A TransportType enum represents the current transport type used by CopySoap web service.
        /// </summary>
        private static TransportProtocol currentTransportType;

        /// <summary>
        /// A bool value represents whether the helper has been initialize.
        /// </summary>
        private static bool hasInitialized = false;

        /// <summary>
        /// An object instance is used for lock blocks which is used for thread safety. This instance is used to keep asynchronous process for visiting the initialization status..  
        /// </summary>
        private static object threadLockObjectOfInitiazeStatus = new object();

        /// <summary>
        /// An int value represents the SOAP time out value in milliseconds
        /// </summary>
        private static int soapTimeOutValue;

        /// <summary>
        /// A string represents the target service URL of MS-COPYS
        /// </summary>
        private static string targetServiceUrlOfMSCOPYS;
        #endregion fields
        
        #region properties

        /// <summary>
        /// Gets the user name of the default user.
        /// </summary>
        public static string DefaultUser
        {
            get
            {
                CheckInitializationStatus();
                return defaultUser;
            }
        }

        /// <summary>
        /// Gets the password of the default user.
        /// </summary>
        public static string PasswordOfDefaultUser
        {
            get
            {
                CheckInitializationStatus();
                return passwordOfDefaultUser;
            }
        }

        /// <summary>
        /// Gets the domain of the default user.
        /// </summary>
        public static string DomainOfDefaultUser
        {
            get
            {
                CheckInitializationStatus();
                return domainOfDefaultUser;
            }
        }

        /// <summary>
        /// Gets the current SOAP version.
        /// </summary>
        public static SoapVersion CurrentSoapVersion
        {
            get
            {
                CheckInitializationStatus();
                return currentSoapVersion;
            }
        }

        /// <summary>
        /// Gets the current transport type.
        /// </summary>
        public static TransportProtocol CurrentTransportType
        {
            get
            {
                CheckInitializationStatus();
                return currentTransportType;
            }
        }

        /// <summary>
        /// Gets the target service URL of MS-COPYS.
        /// </summary>
        public static string TargetServiceUrlOfMSCOPYS
        {
            get
            {
                CheckInitializationStatus();
                return targetServiceUrlOfMSCOPYS;
            }
        }

        /// <summary>
        /// Gets the SOAP time out value in milliseconds.
        /// </summary>
        public static int CurrentSoapTimeOutValue
        {
            get
            {
                CheckInitializationStatus();
                return soapTimeOutValue;
            }
        }

        /// <summary>
        /// Gets the network credentials for the default user.
        /// </summary>
        public static NetworkCredential DefaultUserCredential
        {
            get
            {
                CheckInitializationStatus();
                return new NetworkCredential(defaultUser, passwordOfDefaultUser, domainOfDefaultUser);
            }
        } 

        #endregion properties
        
        /// <summary>
        /// A method used to initialize the helper. It will get necessary values from configuration properties.
        /// </summary>
        /// <param name="testSiteInstance">A parameter represents the ITestSite instance.</param>
        public static void Initialize(ITestSite testSiteInstance)
        {
            if (null == testSiteInstance)
            {
                throw new ArgumentNullException("testSiteInstance");
            }

            lock (threadLockObjectOfInitiazeStatus)
            {
                if (hasInitialized)
                {
                    return;
                }

                testSiteInstance.DefaultProtocolDocShortName = "MS-COPYS";

                // Merge Common configuration file.
                string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSiteInstance);
                Common.MergeGlobalConfig(commonConfigFileName, testSiteInstance);

                // Getting required configuration properties.
                currentSoapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", testSiteInstance);
                currentTransportType = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", testSiteInstance);
                defaultUser = Common.GetConfigurationPropertyValue("UserName", testSiteInstance);
                domainOfDefaultUser = Common.GetConfigurationPropertyValue("Domain", testSiteInstance);
                passwordOfDefaultUser = Common.GetConfigurationPropertyValue("Password", testSiteInstance);
                targetServiceUrlOfMSCOPYS = Common.GetConfigurationPropertyValue("TargetServiceUrlOfDestinationSUT", testSiteInstance);
                int soapTimeOutValueInMintues = Common.GetConfigurationPropertyValue<int>("ServiceTimeOut", testSiteInstance);
                soapTimeOutValue = soapTimeOutValueInMintues * 60000;
 
                // Setting the initialization status.
                hasInitialized = true;
            }
        }

        /// <summary>
        /// A method is used to get SOAP version for proxy class. It is used to for "SoapVersion" property of proxy class.
        /// </summary>
        /// <returns>A return value represents the SOAP version for the test suite.</returns>
        public static SoapProtocolVersion GetSoapProtoclVersionByCurrentSetting()
        {
            CheckInitializationStatus();
            SoapProtocolVersion currentSoapProtocolVersion;
            switch (currentSoapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        currentSoapProtocolVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                case SoapVersion.SOAP12:
                    {
                        currentSoapProtocolVersion = SoapProtocolVersion.Soap12;
                        break;
                    }

                default:
                    {
                        throw new NotSupportedException(string.Format("The test suite dose not support current SOAP version:[{0}]", currentSoapVersion));
                    }
            }

            return currentSoapProtocolVersion;
        }

        /// <summary>
        /// A method used to accept server certificate if current transport type is HTTPS  
        /// </summary>
        public static void AcceptServerCertificate()
        {
            if (TransportProtocol.HTTPS == currentTransportType)
            {
                Common.AcceptServerCertificate();
            }
        }
 
        #region private methods

        /// <summary>
        /// A method used to check the helper whether has been initialized.
        /// </summary>
        private static void CheckInitializationStatus()
        { 
            lock (threadLockObjectOfInitiazeStatus)
            {
               if (!hasInitialized)
               {
                   throw new InvalidOperationException("The TestSuiteManageHelper has not been initialized, call [Initialize] method");
               }
            }
        }

        #endregion private methods
    }
}