namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.ServiceModel;
    using System.ServiceModel.Description;

    /// <summary>
    /// This class is used to create communication channel which will be used by MS-WOPI and MS-FSSHTTP based on the context.
    /// <list type="bullet">
    /// <item>
    ///     <description>The ChannelFactory will be cached and re-used based on the context using the SharedContextCachedEqualityComparer.</description>
    /// </item>
    /// <item>
    ///     <description>The communication channel will not be cached and re-used. Each CreateChannel function call will create a new channel instance.</description>
    /// </item>
    /// <item>
    ///     <description>The created communication channel will be automatically disposed when faults happened.</description>
    /// </item>
    /// </list>
    /// </summary>
    public class SharedChannelManager
    {
        /// <summary>
        /// Specify the lock object for caching the ChannelFactory.
        /// </summary>
        private static object lockObject = new object();

        /// <summary>
        /// Specify the dictionary for caching the ChannelFactory.
        /// </summary>
        private Dictionary<SharedContext, ChannelFactory>
            cached = new Dictionary<SharedContext, ChannelFactory>(new SharedContextCachedEqualityComparer());

        /// <summary>
        /// This method is used to create communication channel based on the specified context.
        /// </summary>
        /// <typeparam name="T">Specify the type of the channel.</typeparam>
        /// <param name="context">Specify the context.</param>
        /// <returns>Return the created new channel instance.</returns>
        public T CreateChannel<T>(SharedContext context)
                where T : IClientChannel
        {
            if (!this.cached.ContainsKey(context))
            {
                lock (lockObject)
                {
                    if (!this.cached.ContainsKey(context))
                    {
                        ChannelFactory<T> cachedItem = new ChannelFactory<T>(context.EndpointConfigurationName, new EndpointAddress(context.TargetUrl));
                        
                        // Try to get the encoding.
                        WOPIMessageEncodingBindingElement messageEncodingBinding = cachedItem.Endpoint.Binding.CreateBindingElements().OfType<WOPIMessageEncodingBindingElement>().FirstOrDefault();
                        if (messageEncodingBinding != null && messageEncodingBinding.Encoding != null)
                        {
                            context.AddOrUpdate("Encoding", messageEncodingBinding.Encoding);
                        }

                        cachedItem.Endpoint.Behaviors.Add(new MessageInspector(context));

                        // Remove the possible visual studio wcf debug behaviors.
                        var removeDiagnosticsEndPoint = cachedItem.Endpoint.Behaviors.Where(behavior => string.Compare("Microsoft.VisualStudio.Diagnostics.ServiceModelSink.Behavior", behavior.GetType().FullName, StringComparison.OrdinalIgnoreCase) == 0).FirstOrDefault();
                        if (removeDiagnosticsEndPoint != null)
                        {
                            cachedItem.Endpoint.Behaviors.Remove(removeDiagnosticsEndPoint);
                        }

                        // Change the credential
                        this.ChangeCredential(cachedItem, context);

                        cachedItem.Faulted += new EventHandler(this.ChannelFactory_Faulted);
                        cachedItem.Open();
                        this.cached.Add(context, cachedItem);
                    }
                }
            }

            ChannelFactory<T> channelFactory = this.cached[context] as ChannelFactory<T>;
            T channel = channelFactory.CreateChannel();
            channel.Faulted += new EventHandler(this.Channel_Faulted);
            channel.Open();
            return channel;
        }

        /// <summary>
        /// This method is used to change the credential for the channel factory.
        /// </summary>
        /// <param name="factory">Specify the ChannelFactory which needs modification of credential.</param>
        /// <param name="context">Specify the shared context.</param>
        private void ChangeCredential(ChannelFactory factory, SharedContext context)
        {
            var defaultCredentials = factory.Endpoint.Behaviors.Find<ClientCredentials>();
            if (defaultCredentials != null)
            {
                factory.Endpoint.Behaviors.Remove(defaultCredentials);
            }

            ClientCredentials credentials = new ClientCredentials();
            credentials.Windows.AllowedImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Impersonation;
            credentials.Windows.ClientCredential.UserName = context.UserName;
            credentials.Windows.ClientCredential.Password = context.Password;
            credentials.Windows.ClientCredential.Domain = context.Domain;

            factory.Endpoint.Behaviors.Add(credentials); 
        }

        /// <summary>
        /// Callback function will be used when fault happened using the channel.
        /// </summary>
        /// <param name="sender">Specify the callback event sender.</param>
        /// <param name="e">Specify the event argument.</param>
        private void Channel_Faulted(object sender, EventArgs e)
        {
            IClientChannel channel = sender as IClientChannel;
            channel.Abort();
            ((IDisposable)channel).Dispose();
        }

        /// <summary>
        /// Callback function will be used when fault happened using the channel factory.
        /// </summary>
        /// <param name="sender">Specify the callback event sender.</param>
        /// <param name="e">Specify the event argument.</param>
        private void ChannelFactory_Faulted(object sender, EventArgs e)
        {
            ChannelFactory factory = (ChannelFactory)sender;
            
            // No matter anything happened, abort this.
            factory.Abort();

            lock (lockObject)
            {
                SharedContext key = null;
                foreach (KeyValuePair<SharedContext, ChannelFactory> pair in this.cached)
                {
                    if (pair.Value == factory)
                    {
                        key = pair.Key;
                        break;
                    }
                }

                if (key != null)
                {
                    this.cached.Remove(key);
                }
            }
        }

        /// <summary>
        /// This class is used to compare the context when caching the ChannelFactory.
        /// </summary>
        public class SharedContextCachedEqualityComparer : IEqualityComparer<SharedContext>
        {
            /// <summary>
            ///  Override to determine whether the specified context are equal.
            /// </summary>
            /// <param name="x">The first context.</param>
            /// <param name="y">The second context.</param>
            /// <returns>Return true if the specified contexts are equal; otherwise, false.</returns>
            public bool Equals(SharedContext x, SharedContext y)
            {
                if (x == null && y == null)
                {
                    return true;
                }

                if (x == null || y == null)
                {
                    return false;
                }

                return x.EndpointConfigurationName == y.EndpointConfigurationName
                    && x.TargetUrl == y.TargetUrl
                    && x.UserName == y.UserName
                    && x.Password == y.Password
                    && x.Domain == y.Domain;
            }

            /// <summary>
            /// Calculate a hash code for the specified context.
            /// </summary>
            /// <param name="obj">Specify the context.</param>
            /// <returns>Return a hash code for the specified object.</returns>
            public int GetHashCode(SharedContext obj)
            {
                int ret = obj.EndpointConfigurationName.GetHashCode();
                ret ^= obj.TargetUrl.GetHashCode();
                ret ^= obj.UserName.GetHashCode();
                ret ^= obj.Password.GetHashCode();
                ret ^= obj.Domain.GetHashCode();

                return ret;
            }
        }
    }
}