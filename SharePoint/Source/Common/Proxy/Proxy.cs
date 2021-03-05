namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Emit;
    using System.Threading;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class that use to dynamic create proxy.
    /// </summary>
    public static class Proxy
    {
        /// <summary>
        /// Specify the lock object for caching the type.
        /// </summary>
        private static object lockObject = new object();

        /// <summary>
        /// Specify the dictionary for caching the type.
        /// </summary>
        private static Dictionary<string, Type> cachedTypes = new Dictionary<string, Type>();

        /// <summary>
        /// Create a new proxy base the specify type.
        /// </summary>
        /// <typeparam name="T">The specify type.</typeparam>
        /// <param name="testSite">Transfer ITestSite into this class, make this class can use ITestSite's function.</param>
        /// <param name="throwException">Indicate that whether throw exception when schema validation not success.</param>
        /// <param name="performSchemaValidation">Indicate that whether perform schema validation.</param>
        /// <param name="ignoreSoapFaultSchemaValidationForSoap12">Indicate that whether ignore schema validation for SOAP fault in SOAP1.2 .</param>
        /// <returns>The new proxy.</returns>
        public static T CreateProxy<T>(ITestSite testSite, bool throwException = false, bool performSchemaValidation = true, bool ignoreSoapFaultSchemaValidationForSoap12 = false)
        where T : WebClientProtocol, new()
        {
            string typeFullName = typeof(T).FullName;
            if (!cachedTypes.ContainsKey(typeFullName))
            {
                lock (lockObject)
                {
                    if (!cachedTypes.ContainsKey(typeFullName))
                    {
                        Type proxyType = CreateType(typeof(T), testSite);
                        cachedTypes.Add(typeFullName, proxyType);
                    }
                }
            }

            Type type = cachedTypes[typeFullName];
            T ret = Activator.CreateInstance(type) as T;
            type.GetMethod("add_After").Invoke(ret, new object[] { new EventHandler<CustomerEventArgs>(new ValidateUtil(testSite, throwException, performSchemaValidation, ignoreSoapFaultSchemaValidationForSoap12).ValidateSchema) });
            return ret;
        }

        /// <summary>
        /// Get parameters type from ParameterInfo.
        /// </summary>
        /// <param name="ps">The array of ParameterInfo.</param>
        /// <returns>The type array of parameters.</returns>
        private static Type[] ParameterInfoToTypeInfo(ParameterInfo[] ps)
        {
            Type[] ts = new Type[ps.Length];

            for (int i = 0; i < ps.Length; i++)
            {
                ts[i] = ps[i].ParameterType;
            }

            return ts;
        }

        /// <summary>
        /// Create a new type base the input type.
        /// </summary>
        /// <param name="baseType">The base type that inherited by new type.</param>
        /// <param name="site">Transfer ITestSite into this operation.</param>
        /// <returns>The new type.</returns>
        private static Type CreateType(Type baseType, ITestSite site)
        {
            string guid = Guid.NewGuid().ToString().Replace("-", string.Empty);
            AssemblyName name = new AssemblyName();
            name.Name = string.Format("{0}{1}", site.DefaultProtocolDocShortName, guid);
            AssemblyBuilder assembly = Thread.GetDomain().DefineDynamicAssembly(name, AssemblyBuilderAccess.Run);
            ModuleBuilder module = assembly.DefineDynamicModule(string.Format("{0}.dll", name.Name), true);
            TypeBuilder type = module.DefineType(
                string.Format("{0}{1}Proxy", site.DefaultProtocolDocShortName, guid),
                TypeAttributes.Public,
                baseType);

            MethodInfo[] abeforerinterlocked = typeof(Interlocked).GetMethods();
            MethodInfo compareexchange = abeforerinterlocked.Where(m => m.IsGenericMethodDefinition == true && m.Name == "CompareExchange").FirstOrDefault();
            MethodInfo compareexchangeinfo = compareexchange.MakeGenericMethod(typeof(System.EventHandler<CustomerEventArgs>));
            FieldBuilder fieldInjector = type.DefineField("fieldInjector", typeof(XmlWriterInjector), FieldAttributes.Private); // Build Event
            Type afterEventType = typeof(EventHandler<CustomerEventArgs>);
            FieldBuilder fieldAfter = type.DefineField("After", afterEventType, FieldAttributes.Public);
            EventBuilder afterEvent = type.DefineEvent("After", EventAttributes.None, afterEventType);
            MethodBuilder afterEventAdd = type.DefineMethod("add_After", MethodAttributes.Public, null, new Type[] { afterEventType });
            afterEventAdd.SetImplementationFlags(MethodImplAttributes.Managed);
            ILGenerator afterILGenerator = afterEventAdd.GetILGenerator();
            afterILGenerator.DeclareLocal(typeof(System.EventHandler<CustomerEventArgs>));
            afterILGenerator.DeclareLocal(typeof(System.EventHandler<CustomerEventArgs>));
            afterILGenerator.DeclareLocal(typeof(System.EventHandler<CustomerEventArgs>));
            afterILGenerator.DeclareLocal(typeof(bool));
            Label lbafter = afterILGenerator.DefineLabel();
            afterILGenerator.Emit(OpCodes.Ldarg_0);
            afterILGenerator.Emit(OpCodes.Ldfld, fieldAfter);
            afterILGenerator.Emit(OpCodes.Stloc_0);
            afterILGenerator.MarkLabel(lbafter);
            afterILGenerator.Emit(OpCodes.Ldloc_0);
            afterILGenerator.Emit(OpCodes.Stloc_1);
            afterILGenerator.Emit(OpCodes.Ldloc_1);
            afterILGenerator.Emit(OpCodes.Ldarg_1);
            afterILGenerator.Emit(OpCodes.Call, typeof(Delegate).GetMethod("Combine", new Type[] { afterEventType, afterEventType }));
            afterILGenerator.Emit(OpCodes.Castclass, afterEventType);
            afterILGenerator.Emit(OpCodes.Stloc_2);
            afterILGenerator.Emit(OpCodes.Ldarg_0);
            afterILGenerator.Emit(OpCodes.Ldflda, fieldAfter);
            afterILGenerator.Emit(OpCodes.Ldloc_2);
            afterILGenerator.Emit(OpCodes.Ldloc_1);
            afterILGenerator.Emit(OpCodes.Call, compareexchangeinfo);
            afterILGenerator.Emit(OpCodes.Stloc_0);
            afterILGenerator.Emit(OpCodes.Ldloc_0);
            afterILGenerator.Emit(OpCodes.Ldloc_1);
            afterILGenerator.Emit(OpCodes.Ceq);
            afterILGenerator.Emit(OpCodes.Ldc_I4_0);
            afterILGenerator.Emit(OpCodes.Ceq);
            afterILGenerator.Emit(OpCodes.Stloc_3);
            afterILGenerator.Emit(OpCodes.Ldloc_3);
            afterILGenerator.Emit(OpCodes.Brtrue_S, lbafter);
            afterILGenerator.Emit(OpCodes.Ret);
            afterEvent.SetAddOnMethod(afterEventAdd);

            // Try to override the virtual function here
            MethodInfo mi = baseType.GetMethod("GetReaderForMessage", BindingFlags.NonPublic | BindingFlags.Instance);
            MethodBuilder mb = type.DefineMethod(
                mi.Name,
                MethodAttributes.Public | MethodAttributes.Virtual,
                mi.CallingConvention,
                mi.ReturnType,
                ParameterInfoToTypeInfo(mi.GetParameters()));
            type.DefineMethodOverride(mb, mi);
            ILGenerator ilgenerator = mb.GetILGenerator();
            Type objCustomerEventArgs = typeof(CustomerEventArgs);
            ilgenerator.DeclareLocal(objCustomerEventArgs);
            ilgenerator.DeclareLocal(typeof(System.Xml.XmlReader));
            ilgenerator.DeclareLocal(typeof(bool));
            Label lbxmlreader = ilgenerator.DefineLabel();
            Label lbbool = ilgenerator.DefineLabel();
            ConstructorInfo objReaderInfo = objCustomerEventArgs.GetConstructor(new Type[0]);
            ilgenerator.Emit(OpCodes.Nop);
            ilgenerator.Emit(OpCodes.Ldarg_0);
            ilgenerator.Emit(OpCodes.Ldfld, fieldAfter);
            ilgenerator.Emit(OpCodes.Ldnull);
            ilgenerator.Emit(OpCodes.Ceq);
            ilgenerator.Emit(OpCodes.Stloc_2);
            ilgenerator.Emit(OpCodes.Ldloc_2);
            ilgenerator.Emit(OpCodes.Brtrue_S, lbbool);
            ilgenerator.Emit(OpCodes.Nop);
            ilgenerator.Emit(OpCodes.Newobj, objReaderInfo);
            ilgenerator.Emit(OpCodes.Stloc_0);
            ilgenerator.Emit(OpCodes.Ldloc_0);
            ilgenerator.Emit(OpCodes.Ldarg_1);
            ilgenerator.Emit(OpCodes.Callvirt, objCustomerEventArgs.GetMethod("set_ResponseMessage", new Type[] { typeof(System.Web.Services.Protocols.SoapClientMessage) }));
            ilgenerator.Emit(OpCodes.Nop);
            ilgenerator.Emit(OpCodes.Ldloc_0);
            ilgenerator.Emit(OpCodes.Ldarg_0);
            ilgenerator.Emit(OpCodes.Ldfld, fieldInjector);
            ilgenerator.Emit(OpCodes.Callvirt, objCustomerEventArgs.GetMethod("set_RequestMessage", new Type[] { typeof(XmlWriterInjector) }));
            ilgenerator.Emit(OpCodes.Nop);
            ilgenerator.Emit(OpCodes.Ldarg_0);
            ilgenerator.Emit(OpCodes.Ldfld, fieldAfter);
            ilgenerator.Emit(OpCodes.Ldarg_0);
            ilgenerator.Emit(OpCodes.Ldloc_0);
            ilgenerator.Emit(OpCodes.Callvirt, afterEventType.GetMethod("Invoke"));
            ilgenerator.Emit(OpCodes.Nop);
            ilgenerator.Emit(OpCodes.Ldloc_0);
            ilgenerator.Emit(OpCodes.Callvirt, objCustomerEventArgs.GetMethod("get_ValidationXmlReaderOut"));
            ilgenerator.Emit(OpCodes.Stloc_1);
            ilgenerator.Emit(OpCodes.Br_S, lbxmlreader);
            ilgenerator.MarkLabel(lbbool);
            ilgenerator.Emit(OpCodes.Ldarg_0);
            ilgenerator.Emit(OpCodes.Ldarg_1);
            ilgenerator.Emit(OpCodes.Ldarg_2);
            ilgenerator.Emit(OpCodes.Call, baseType.GetMethod("GetReaderForMessage", BindingFlags.NonPublic | BindingFlags.Instance));
            ilgenerator.Emit(OpCodes.Stloc_1);
            ilgenerator.Emit(OpCodes.Br_S, lbxmlreader);
            ilgenerator.MarkLabel(lbxmlreader);
            ilgenerator.Emit(OpCodes.Ldloc_1);
            ilgenerator.Emit(OpCodes.Ret);

            MethodInfo miwriter = baseType.GetMethod("GetWriterForMessage", BindingFlags.NonPublic | BindingFlags.Instance);
            MethodBuilder mbwriter = type.DefineMethod(
                miwriter.Name,
                MethodAttributes.Public | MethodAttributes.Virtual,
                miwriter.CallingConvention,
                miwriter.ReturnType,
                ParameterInfoToTypeInfo(miwriter.GetParameters()));
            type.DefineMethodOverride(mbwriter, miwriter);
            ILGenerator ilgwriter = mbwriter.GetILGenerator();
            Type objxmlwriterinjector = typeof(XmlWriterInjector);
            ilgwriter.DeclareLocal(typeof(XmlWriter));
            Label returnwriterLabel = ilgwriter.DefineLabel();
            ConstructorInfo objxmlwriterinjectorinfo = objxmlwriterinjector.GetConstructor(new Type[] { typeof(XmlWriter) });
            ilgwriter.Emit(OpCodes.Ldarg_0);
            ilgwriter.Emit(OpCodes.Ldarg_0);
            ilgwriter.Emit(OpCodes.Ldarg_1);
            ilgwriter.Emit(OpCodes.Ldarg_2);
            ilgwriter.Emit(OpCodes.Call, baseType.GetMethod("GetWriterForMessage", BindingFlags.NonPublic | BindingFlags.Instance));
            ilgwriter.Emit(OpCodes.Newobj, objxmlwriterinjectorinfo);
            ilgwriter.Emit(OpCodes.Stfld, fieldInjector);
            ilgwriter.Emit(OpCodes.Ldarg_0);
            ilgwriter.Emit(OpCodes.Ldfld, fieldInjector);
            ilgwriter.Emit(OpCodes.Stloc_0);
            ilgwriter.Emit(OpCodes.Br_S, returnwriterLabel);
            ilgwriter.MarkLabel(returnwriterLabel);
            ilgwriter.Emit(OpCodes.Ldloc_0);
            ilgwriter.Emit(OpCodes.Ret);
            Type returntype = type.CreateType();
            return returntype;
        }
    }
}