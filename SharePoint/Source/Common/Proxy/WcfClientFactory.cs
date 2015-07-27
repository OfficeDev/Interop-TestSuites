//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.ServiceModel;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This class is used to create WCF client.
    /// </summary>
    public static class WcfClientFactory
    {
        /// <summary>
        /// Create a WCF client based on the specified type T to add the schema validation ability.
        /// </summary>
        /// <typeparam name="T">Specify the WCF client type which is derived class ClientBase<![CDATA[<I>]]> </typeparam>
        /// <typeparam name="I">Specify the WCF contract interface.</typeparam>
        /// <param name="site">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="endPointConfigurationName">Specify the endpoint configuration name.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2 .</param>
        /// <returns>Return the new WCF client.</returns>
        public static T CreateClient<T, I>(
                                ITestSite site,
                                string endPointConfigurationName,
                                bool throwException = false,
                                bool performSchemaValidation = true,
                                bool ignoreSoapFaultSchemaValidationForSoap12 = false)
            where T : ClientBase<I>
            where I : class
        {
            T ret = Activator.CreateInstance(typeof(T), new object[] { endPointConfigurationName }) as T;
            ResponseSchemaValidationInspector messageInjector = new ResponseSchemaValidationInspector();
            messageInjector.ValidationEvent += new EventHandler<CustomerEventArgs>(new ValidateUtil(site, throwException, performSchemaValidation, ignoreSoapFaultSchemaValidationForSoap12).ValidateSchema);
            ((ClientBase<I>)ret).Endpoint.Behaviors.Add(messageInjector);

            return ret;
        }

        /// <summary>
        /// Create a WCF client based on the specified type T to add the schema validation ability.
        /// </summary>
        /// <typeparam name="T">Specify the WCF client type which is derived class ClientBase<![CDATA[<I>]]> </typeparam>
        /// <typeparam name="I">Specify the WCF contract interface.</typeparam>
        /// <param name="site">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="endPointConfigurationName">Specify the endpoint configuration name.</param>
        /// <param name="remoteAddress">Specify the remote address.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2 .</param>
        /// <returns>Return the new WCF client.</returns>
        public static T CreateClient<T, I>(
                                ITestSite site,
                                string endPointConfigurationName,
                                string remoteAddress,
                                bool throwException = false,
                                bool performSchemaValidation = true,
                                bool ignoreSoapFaultSchemaValidationForSoap12 = false)
            where T : ClientBase<I>
            where I : class
        {
            T ret = Activator.CreateInstance(typeof(T), new object[] { endPointConfigurationName, remoteAddress }) as T;
            ResponseSchemaValidationInspector messageInjector = new ResponseSchemaValidationInspector();
            messageInjector.ValidationEvent += new EventHandler<CustomerEventArgs>(new ValidateUtil(site, throwException, performSchemaValidation, ignoreSoapFaultSchemaValidationForSoap12).ValidateSchema);
            ((ClientBase<I>)ret).Endpoint.Behaviors.Add(messageInjector);

            return ret;
        }

        /// <summary>
        /// Create a WCF client based on the specified type T to add the schema validation ability.
        /// </summary>
        /// <typeparam name="T">Specify the WCF client type which is derived class ClientBase<![CDATA[<I>]]> </typeparam>
        /// <typeparam name="I">Specify the WCF contract interface.</typeparam>
        /// <param name="site">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="endPointConfigurationName">Specify the endpoint configuration name.</param>
        /// <param name="remoteAddress">Specify the remote address.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2 .</param>
        /// <returns>Return the new WCF client.</returns>
        public static T CreateClient<T, I>(
                                ITestSite site,
                                string endPointConfigurationName,
                                EndpointAddress remoteAddress,
                                bool throwException = false,
                                bool performSchemaValidation = true,
                                bool ignoreSoapFaultSchemaValidationForSoap12 = false)
            where T : ClientBase<I>
            where I : class
        {
            T ret = Activator.CreateInstance(typeof(T), new object[] { endPointConfigurationName, remoteAddress }) as T;
            ResponseSchemaValidationInspector messageInjector = new ResponseSchemaValidationInspector();
            messageInjector.ValidationEvent += new EventHandler<CustomerEventArgs>(new ValidateUtil(site, throwException, performSchemaValidation, ignoreSoapFaultSchemaValidationForSoap12).ValidateSchema);
            ((ClientBase<I>)ret).Endpoint.Behaviors.Add(messageInjector);

            return ret;
        }

        /// <summary>
        /// Create a WCF client based on the specified type T to add the schema validation ability.
        /// </summary>
        /// <typeparam name="T">Specify the WCF client type which is derived class ClientBase<![CDATA[<I>]]> </typeparam>
        /// <typeparam name="I">Specify the WCF contract interface.</typeparam>
        /// <param name="site">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="binding">Specify the WCF binding.</param>
        /// <param name="remoteAddress">Specify the remote address.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2 .</param>
        /// <returns>Return the new WCF client.</returns>
        public static T CreateClient<T, I>(
                                ITestSite site,
                                System.ServiceModel.Channels.Binding binding,
                                EndpointAddress remoteAddress,
                                bool throwException = false,
                                bool performSchemaValidation = true,
                                bool ignoreSoapFaultSchemaValidationForSoap12 = false)
            where T : ClientBase<I>
            where I : class
        {
            T ret = Activator.CreateInstance(typeof(T), new object[] { binding, remoteAddress }) as T;
            ResponseSchemaValidationInspector messageInjector = new ResponseSchemaValidationInspector();
            messageInjector.ValidationEvent += new EventHandler<CustomerEventArgs>(new ValidateUtil(site, throwException, performSchemaValidation, ignoreSoapFaultSchemaValidationForSoap12).ValidateSchema);
            ((ClientBase<I>)ret).Endpoint.Behaviors.Add(messageInjector);

            return ret;
        }
    }
}