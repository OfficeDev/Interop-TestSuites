namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Net;
    using System.Text;
    using System.Threading;
    using System.Xml;

    /// <summary>
    /// This class is used to help the implementation of the discovery operation.
    /// </summary>
    public class DiscoveryRequestListener : HelperBase, IDisposable
    {
        /// <summary>
        /// A bool value indicating whether the listen thread has been started by one instance of this type. 
        /// </summary>
        private static bool hasStartListenThread = false;

        /// <summary>
        /// A bool value indicating whether the listen thread has response a discovery request succeed. 
        /// </summary>
        private static bool hasResponseDiscoveryRequestSucceed = false;

        /// <summary>
        /// A thread handle indicating the instance of the listen thread.
        /// </summary>
        private static Thread listenThreadHandle = null;

        /// <summary>
        /// An object instance is used for lock blocks which is used for multiple threads. This instance is used to keep asynchronous process for verifying whether the listen thread has been started.
        /// </summary>
        private static object threadLockStaticObjectForVisitThread = new object();

        /// <summary>
        /// An object instance is used for lock blocks which is used for multiple threads. This instance is used to keep asynchronous process for append log in different threads.  
        /// </summary>
        private static object threadLockObjectForAppendLog = new object();

        /// <summary>
        /// A Type instance represents the current helper's type information.
        /// </summary>
        private static Type currentHelperType;

        /// <summary>
        /// Initializes a new instance of the <see cref="DiscoveryRequestListener"/> class.
        /// </summary>
        /// <param name="hostDiscoveryMachineName">A parameter represents the machine name which will listen the discovery request. The value must be the name of the current machine.</param>
        /// <param name="responseXmlForDiscovery">A parameter represents the discovery response which will response to WOPI server.</param>
        public DiscoveryRequestListener(string hostDiscoveryMachineName, string responseXmlForDiscovery)
        {
            if (string.IsNullOrEmpty(hostDiscoveryMachineName))
            {
                throw new ArgumentNullException("hostDiscoveryMachineName");
            }

            if (string.IsNullOrEmpty(responseXmlForDiscovery))
            {
                throw new ArgumentNullException("responseXmlForDiscovery");
            }

            if (null == currentHelperType)
            {
                currentHelperType = this.GetType();
            }

            this.HostNameOfDiscoveryService = hostDiscoveryMachineName;
            this.ResponseDiscovery = responseXmlForDiscovery;
            if (null == ListenInstance)
            {
                ListenInstance = new HttpListener();
            }
           
            this.IsRequiredStop = false;
            this.IsDisposed = false;
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="DiscoveryRequestListener"/> class. This method will be invoked by .net GC collector automatically.
        /// </summary>
        ~DiscoveryRequestListener()
        {  
            lock (threadLockStaticObjectForVisitThread)
            {
                this.Dispose(false);
            }
        }
 
        #region properties

        /// <summary>
        /// Gets a value indicating whether the DiscoveryRequestListener has responded to a discovery request successfully.
        /// </summary>
        public static bool HasResponseSucceed
        {
            get
            {
                lock (threadLockStaticObjectForVisitThread)
                {
                    return hasResponseDiscoveryRequestSucceed;
                }
            }
        }

        /// <summary>
        /// Gets or sets the HttpListener type instance.
        /// </summary>
        protected static HttpListener ListenInstance { get; set; }

        /// <summary>
        /// Gets or sets the host name which will listen and response for the discovery request.
        /// </summary>
        protected string HostNameOfDiscoveryService { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the DiscoveryRequestListener type has released related resource.
        /// </summary>
        protected bool IsDisposed { get; set; }

        /// <summary>
        /// Gets or sets the response information for the discovery request.
        /// </summary>
        protected string ResponseDiscovery { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the listening thread is required to stop.
        /// </summary>
        protected bool IsRequiredStop { get; set; }

        #endregion 

        /// <summary>
        /// A method is used to implement the IDisposable interface, it allows the user to dispose the current instance if user need to release allocated resources.
        /// </summary>
        public void Dispose()
        {
            lock (threadLockStaticObjectForVisitThread)
            {
                this.Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        /// <summary>
        /// A method is used to start the listen thread to listen the discovery request.
        /// </summary>
        /// <returns>A return value represents the thread instance handle, which is processing the listen logic. This thread instance can be used to control the thread's status and clean up.</returns>
        public Thread StartListen()
        {
            // Verify whether the listen thread has been started from a DiscoveryRequestListener type instance.
            lock (threadLockStaticObjectForVisitThread)
            {
                lock (threadLockObjectForAppendLog)
                {
                    DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, @"Try to start listener thread from current thread.");
                }

                if (null == ListenInstance)
                {
                    ListenInstance = new HttpListener();
                }

                if (hasStartListenThread)
                {
                    lock (threadLockObjectForAppendLog)
                    {
                        DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format(@"The listen thread [{0}] exists.", listenThreadHandle.ManagedThreadId));
                    }

                    return listenThreadHandle;
                }

                if (!ListenInstance.IsListening)
                {
                    this.SetPrefixForListener(ListenInstance);
                }

                while (!ListenInstance.IsListening)
                {
                    // Sleep 1 second to wait until the status is stable, and allow other threads get the control to update the status.
                    Thread.Sleep(1000);
                }

                listenThreadHandle = new Thread(this.ListenToRequest);
                listenThreadHandle.Name = "Listen Discovery request thread";
                listenThreadHandle.Start();

                lock (threadLockObjectForAppendLog)
                {
                    DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Start the listening thread. The listening thread managed Id[{0}]", listenThreadHandle.ManagedThreadId));
                }

                // Set the status to indicate there has started a listen thread.
                hasStartListenThread = true;
                return listenThreadHandle;
            }
        }

        /// <summary>
        /// A method is used to stop listen process. This method will abort the thread which is listening discovery request and release all resource are used by the thread.
        /// </summary>
        public void StopListen()
        {
            lock (threadLockStaticObjectForVisitThread)
            {   
                // If the listen thread has not been start, skip the stop operation.
                if (!hasStartListenThread)
                {
                    return;
                }

                this.IsRequiredStop = true;

                lock (threadLockObjectForAppendLog)
                {
                    DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Stop the listening thread.The listening thread managed Id[{0}]", listenThreadHandle.ManagedThreadId));
                }

                if (listenThreadHandle != null && listenThreadHandle.ThreadState != ThreadState.Unstarted
                   && ListenInstance != null)
                {
                    lock (threadLockObjectForAppendLog)
                    {
                        // Close the http listener and release the resource used by listener. This might cause the thread generate exception and then the thread will be expected to end and join to the main thread.
                        ListenInstance.Abort();
                        ((IDisposable)ListenInstance).Dispose();
                        DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Release the Httplistener resource. The listening thread managed Id[{0}]", listenThreadHandle.ManagedThreadId));
                    }

                    // Wait the thread join to the main caller thread.
                    TimeSpan listenThreadJoinTimeOut = new TimeSpan(0, 0, 1);
                    bool isthreadEnd = listenThreadHandle.Join(listenThreadJoinTimeOut);

                    // If the thread could not end as expected, abort this thread.
                    if (!isthreadEnd)
                    {
                        if ((listenThreadHandle.ThreadState & (ThreadState.Stopped | ThreadState.Unstarted)) == 0)
                        {
                            listenThreadHandle.Abort();
                            lock (threadLockObjectForAppendLog)
                            {
                                DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Abort the listening thread. The listening thread managed Id[{0}]", listenThreadHandle.ManagedThreadId));
                            }
                        }
                    }

                    // Set the static status to tell other instance, the listen thread has been abort.
                    hasStartListenThread = false;
                    listenThreadHandle = null;
                }
            }
        }

        #region protected method

        /// <summary>
        /// A method is used to perform custom dispose logic when the GC try to collect this instance.
        /// </summary>
        /// <param name="disposing">A parameter represents the disposing way, the 'true' means it is called from user code by calling IDisposable.Dispose, otherwise it means the GC is trying to process this instance.</param>
        protected virtual void Dispose(bool disposing)
        {
            lock (threadLockStaticObjectForVisitThread)
            {
                if (!this.IsDisposed)
                {
                    this.StopListen();

                    if (disposing)
                    {
                        ListenInstance = null;
                    }

                    this.IsDisposed = true;
                }
            }
        }

        /// <summary>
        /// A method is used to listening the discovery request. It will be executed by a thread which is started on ListenThreadInstance method.
        /// </summary>
        protected void ListenToRequest()
        {
            // If the listener is listening, just keep on execute below code.
            while (ListenInstance.IsListening)
            {
                // if the calling thread requires stopping the listening mission, just return and exit the loop. This value of "IsrequireStop" property is managed by "StopListen" method.
                if (this.IsRequiredStop)
                {
                    break;
                }

                lock (threadLockStaticObjectForVisitThread)
                {
                    // Double check the "IsrequireStop" status.
                    if (this.IsRequiredStop)
                    {
                        break;
                    }
                }

                lock (threadLockObjectForAppendLog)
                {
                    string logMsg = string.Format("Listening............ The listen thread: managed id[{0}].", Thread.CurrentThread.ManagedThreadId);
                    DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, logMsg);
                }

                // Get an incoming request.
                HttpListenerContext listenContext = null;
                HttpListenerResponse responseOfCurrentRequest = null;
                HttpListenerRequest currentRequest = null;

                try
                {
                    listenContext = ListenInstance.GetContext();
                    currentRequest = listenContext.Request as HttpListenerRequest;
                    responseOfCurrentRequest = listenContext.Response as HttpListenerResponse;

                    if (!currentRequest.RawUrl.Equals(@"/hosting/discovery", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                }
                catch (Exception ex)
                {
                    lock (threadLockObjectForAppendLog)
                    {
                        DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("The listen thread catches an [{0}] exception:[{1}].", ex.GetType().Name, ex.Message));
                    }

                    lock (threadLockStaticObjectForVisitThread)
                    {
                        if (this.IsRequiredStop)
                        {
                            lock (threadLockObjectForAppendLog)
                            {
                                DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, "Requires stopping the Httplistener.");
                            }

                            return;
                        }
                        else
                        {
                            this.RestartListener();
                        }
                    }
                }

                if (responseOfCurrentRequest != null)
                {
                    responseOfCurrentRequest.ContentType = "text/xml";
                    responseOfCurrentRequest.ContentEncoding = Encoding.UTF8;
                    responseOfCurrentRequest.StatusCode = 200;

                    // Get the xml response.
                    XmlDocument reponseXml = new XmlDocument();
                    reponseXml.LoadXml(this.ResponseDiscovery);
                    bool writeResponseSucceed = false;
                    try
                    {
                        using (XmlTextWriter xmlWriter = new XmlTextWriter(responseOfCurrentRequest.OutputStream, Encoding.UTF8))
                        {
                            reponseXml.Save(xmlWriter);
                            writeResponseSucceed = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        lock (threadLockObjectForAppendLog)
                        {
                            DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("The listen thread catches an [{0}] exception:[{1}] on responding.", ex.GetType().Name, ex.Message));
                        }

                        lock (threadLockStaticObjectForVisitThread)
                        {
                            if (this.IsRequiredStop)
                            {
                                lock (threadLockObjectForAppendLog)
                                {
                                    DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Catch an exception:[{0}]. Current requires stopping the Httplistener. Thread managed Id[{1}].", ex.Message, Thread.CurrentThread.ManagedThreadId));
                                }

                                return;
                            }
                            else
                            {
                                this.RestartListener();
                            }
                        }
                    }

                    if (writeResponseSucceed)
                    {
                        lock (threadLockStaticObjectForVisitThread)
                        {   
                            // Setting the status.
                            if (!hasResponseDiscoveryRequestSucceed)
                            {
                                hasResponseDiscoveryRequestSucceed = true;
                            }
                        }

                        lock (threadLockObjectForAppendLog)
                        {
                            DiscoveryProcessHelper.AppendLogs(
                                      currentHelperType,
                                      DateTime.Now,
                                      string.Format(
                                                "Response the discovery request from [{0}] succeed! The listen thread managedId[{1}]",
                                                 currentRequest.UserHostName,
                                                 Thread.CurrentThread.ManagedThreadId));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// A method is used to set the prefix for listener.
        /// </summary>
        /// <param name="listenerInstance">A parameter represents the HttpListener instance which will be set the listened prefix.</param>
        protected void SetPrefixForListener(HttpListener listenerInstance)
        {
            if (null == listenerInstance)
            {
                throw new ArgumentNullException("listenerInstance");
            }

            if (listenerInstance.IsListening)
            {
                throw new InvalidOperationException("The listener has been started, could not set listened prefix for it when it is in started status.");
            }

            // The discovery request must send to "http://{0}/hosting/discovery" URL, and the listener's prefix should append the "/" after the url.
            string listenedprefix = string.Format(@"http://{0}/hosting/discovery/", this.HostNameOfDiscoveryService);
            Uri uri = new Uri(listenedprefix);
            listenedprefix = uri.GetComponents(UriComponents.SchemeAndServer, UriFormat.SafeUnescaped);
            ListenInstance.Prefixes.Add(listenedprefix + "/");
            ListenInstance.Start();
            
            lock (threadLockObjectForAppendLog)
            {
                DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, string.Format("Start the HttpListener [{0}] succeed.", Thread.CurrentThread.ManagedThreadId));
            }
        }

        /// <summary>
        /// A method is used to restart the HTTP listener. It will dispose the original HTTP listener and then re-generate a HTTP listen instance to listen request.
        /// </summary>
        protected void RestartListener()
        {
            lock (threadLockObjectForAppendLog)
            {
                DiscoveryProcessHelper.AppendLogs(currentHelperType, DateTime.Now, "Try to restart the Httplistener.");
            }

            // If it is not in "stop" status, try to restart the Httplistener.
            if (ListenInstance.IsListening)
            {
                ListenInstance.Stop();
            }

            while (ListenInstance.IsListening)
            {
                // sleep zero, so that other thread can get the CPU timespan in a valid 
                Thread.Sleep(0);
            }

            // Release the original HttpListener resource.
            ListenInstance.Abort();
            ((IDisposable)ListenInstance).Dispose();
            ListenInstance = null;

            // Restart a new HttpListener instance.
            ListenInstance = new HttpListener();
            this.SetPrefixForListener(ListenInstance);
        }

        #endregion 
    }
}