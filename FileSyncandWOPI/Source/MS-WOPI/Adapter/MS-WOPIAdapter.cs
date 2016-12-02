namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The implementation of the IMS_WOPIAdapter interface.
    /// </summary>
    public partial class MS_WOPIAdapter : ManagedAdapterBase, IMS_WOPIAdapter
    {   
        /// <summary>
        /// A string represents the password for the default user.
        /// </summary>
        private string defaultPassword;

        /// <summary>
        /// A string represents the domain name for the default user.
        /// </summary>
        private string defaultDomain;

        /// <summary>
        /// A string represent the user name for the default user.
        /// </summary>
        private string defaultUserName;

        /// <summary>
        /// A TransportProtocol type represents the current transport used by test suite.
        /// </summary>
        private TransportProtocol currentTransport;

        #region Initialize method

        /// <summary>
        /// The Overridden Initialize method, it includes the initialization logic of this adapter.
        /// </summary>
        /// <param name="testSite">The ITestSite member of ManagedAdapterBase</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            this.defaultUserName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            this.defaultPassword = Common.GetConfigurationPropertyValue("Password", this.Site);
            this.defaultDomain = Common.GetConfigurationPropertyValue("Domain", this.Site);

            this.currentTransport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (TransportProtocol.HTTPS == this.currentTransport)
            {
                Common.AcceptServerCertificate();
            }
        }

        #endregion 

        #region protocol operations

        /// <summary>
        /// This method is used to take a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierValue">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for Lock operation.</returns>
        public WOPIHttpResponse Lock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierValue)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            // Setting required headers
            commonHeaders.Add("X-WOPI-Lock", lockIdentifierValue);
            string wopiOverrideValue = this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.Lock);
            commonHeaders.Add("X-WOPI-Override", wopiOverrideValue);

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.Lock);
            this.ValidateLockResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to update a file on the WOPI server.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSize">A parameter represents the size of the request body.</param>
        /// <param name="bodyContents">A parameter represents the body contents of the request.</param>
        /// <param name="lockIdentifierOfFile">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for PutFile operation.</returns>
        public WOPIHttpResponse PutFile(string targetResourceUrl, WebHeaderCollection commonHeaders, int? xwopiSize, byte[] bodyContents, string lockIdentifierOfFile)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            // Setting required headers
            commonHeaders.Add("X-WOPI-Lock", lockIdentifierOfFile);
            string wopiOverrideValue = this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.PutFile);
            commonHeaders.Add("X-WOPI-Override", wopiOverrideValue);

            // Setting optional headers
            if (xwopiSize.HasValue)
            {
                string xwopiSizeValue = xwopiSize.Value.ToString();
                commonHeaders.Add("X-WOPI-Size", xwopiSizeValue);
            }

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, bodyContents, WOPIOperationName.PutFile);
            this.ValidateFileContentCapture();
            this.ValidatePutFileResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to return information about folder and permissions for the user which is determined by "targetResourceUrl" parameter and the "Authorization" header in "commonHeaders" parameter.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSessionContextValue">A parameter represents the value of the session context information.</param>
        /// <returns>A return value represents the http response for CheckFolderInfo operation.</returns>
        public WOPIHttpResponse CheckFolderInfo(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSessionContextValue)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            if (!string.IsNullOrEmpty(xwopiSessionContextValue))
            {
                commonHeaders.Add("X-WOPI-SessionContext", xwopiSessionContextValue);
            }

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.CheckFolderInfo);
            this.ValidateFoldersCapture();
            this.ValidateCheckFolderInfoResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to return the file information including file properties and permissions for the user who is identified by token that is sent in the request.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSessionContextValue">A parameter represents the value of the session context information.</param>
        /// <returns>A return value represents the http response for CheckFileInfo operation.</returns>
        public WOPIHttpResponse CheckFileInfo(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSessionContextValue)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            if (!string.IsNullOrEmpty(xwopiSessionContextValue))
            {
                commonHeaders.Add("X-WOPI-SessionContext", xwopiSessionContextValue);
            }

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.CheckFileInfo);
            this.ValidateFilesCapture();
            this.ValidateCheckFileInfoResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to get a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="maxExpectedSize">A parameter represents the specifying upper bound size of the file being requested.</param>
        /// <returns>A return value represents the http response for GetFile operation.</returns>
        public WOPIHttpResponse GetFile(string targetResourceUrl, WebHeaderCollection commonHeaders, int? maxExpectedSize)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            // Setting the optional headers
            if (maxExpectedSize.HasValue)
            {
                commonHeaders.Add("X-WOPI-MaxExpectedSize", maxExpectedSize.Value.ToString());
            }

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.GetFile);

            this.ValidateFileContentCapture();

            return responseTemp;
        }

        /// <summary>
        /// This method is used to release a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierOfFile">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for UnLock operation.</returns>
        public WOPIHttpResponse UnLock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierOfFile)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-Lock", lockIdentifierOfFile);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.UnLock));

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.UnLock);
            this.ValidateUnLockResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to return the contents of a folder on the WOPI server.
        /// </summary>
        /// <param name="targetResourceUrlOfFloder">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <returns>A return value represents the http response for EnumerateChildren operation.</returns>
        public WOPIHttpResponse EnumerateChildren(string targetResourceUrlOfFloder, WebHeaderCollection commonHeaders)
        {
            WOPIHttpResponse responseTemp = null;

            responseTemp = this.SendWOPIRequest(targetResourceUrlOfFloder, commonHeaders, null, WOPIOperationName.EnumerateChildren);
            this.ValidateEnumerateChildrenResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to create a file on the WOPI server based on the current file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiSuggestedTarget">A parameter represents the file name in order to create a file.</param>
        /// <param name="xwopiRelativeTarget">A parameter represents the file name of the current file.</param>
        /// <param name="bodyContents">A parameter represents the body contents of the request.</param>
        /// <param name="xwopiOverwriteRelativeTarget">A parameter represents the value that specifies whether the host overwrite the file name if it exists.</param>
        /// <param name="xwopiSize">A parameter represents the size of the file.</param>
        /// <returns>A return value represents the http response for PutRelativeFile operation.</returns>
        public WOPIHttpResponse PutRelativeFile(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiSuggestedTarget, string xwopiRelativeTarget, byte[] bodyContents, bool? xwopiOverwriteRelativeTarget, int? xwopiSize)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            // Setting required headers
            string wopiOverrideValue = this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.PutRelativeFile);
            commonHeaders.Add("X-WOPI-Override", wopiOverrideValue);

            // Setting optional headers
            if (!string.IsNullOrEmpty(xwopiSuggestedTarget))
            {
                commonHeaders.Add("X-WOPI-SuggestedTarget", xwopiSuggestedTarget);
            }

            if (xwopiOverwriteRelativeTarget.HasValue)
            {
                string xwopiOverwriteRelativeTargetValue = xwopiOverwriteRelativeTarget.Value.ToString();
                commonHeaders.Add("X-WOPI-OverwriteRelativeTarget", xwopiOverwriteRelativeTargetValue);
            }

            if (!string.IsNullOrEmpty(xwopiRelativeTarget))
            {
                commonHeaders.Add("X-WOPI-RelativeTarget", xwopiRelativeTarget);
            }

            if (xwopiSize.HasValue)
            {
                string xwopiSizeValue = xwopiSize.Value.ToString();
                commonHeaders.Add("X-WOPI-Size", xwopiSizeValue);
            }

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, bodyContents, WOPIOperationName.PutRelativeFile);

            this.ValidateFilesCapture();
            this.ValidatePutRelativeFileResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to release the existing lock for modifying a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierOfRefreshLock">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for RefreshLock operation.</returns>
        public WOPIHttpResponse RefreshLock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierOfRefreshLock)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-Lock", lockIdentifierOfRefreshLock);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.RefreshLock));

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.RefreshLock);

            this.ValidateFilesCapture();
            this.ValidateRefreshLockResponse(responseTemp);

            return responseTemp;
        }

        /// <summary>
        /// This method is used to refresh and retake a lock for editing a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="lockIdentifierValue">A parameter represents the value which is provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <param name="lockIdentifierOldValue">A parameter represents the value which is previously provided by WOPI client that WOPI server used to identify the lock on the file.</param>
        /// <returns>A return value represents the http response for UnlockAndRelock operation.</returns>
        public WOPIHttpResponse UnlockAndRelock(string targetResourceUrl, WebHeaderCollection commonHeaders, string lockIdentifierValue, string lockIdentifierOldValue)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-Lock", lockIdentifierValue);
            commonHeaders.Add("X-WOPI-OldLock", lockIdentifierOldValue);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.UnlockAndRelock));

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.UnlockAndRelock);

            this.ValidateFilesCapture();
            this.ValidateUnlockAndRelockResponse(responseTemp);

            return responseTemp;
        }
 
        /// <summary>
        /// This method is used to delete a file.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <returns>A return value represents the http response for DeleteFile operation.</returns>
        public WOPIHttpResponse DeleteFile(string targetResourceUrl, WebHeaderCollection commonHeaders)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.DeleteFile));

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.DeleteFile);

            this.ValidateFilesCapture();

            return responseTemp;
        }

        /// <summary>
        /// This method is used to get a link to a file though which a user is able to operate on a file in a limited way.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiRestrictedLink">A parameter represents the type of restricted link being request by WOPI client.</param>
        /// <returns>A return value represents the http response for GetRestrictedLink operation.</returns>
        public WOPIHttpResponse GetRestrictedLink(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiRestrictedLink)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-RestrictedLink", xwopiRestrictedLink);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.GetRestrictedLink));

            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.GetRestrictedLink);

            this.ValidateFilesCapture();

            return responseTemp;
        }

        /// <summary>
        /// This method is used to revoke all links to a file through which a number of users are able to operate on a file in a limited way.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiRestrictedLink">A parameter represents the type of restricted link being revoked by WOPI client.</param>
        /// <returns>A return value represents the http response for RevokeRestrictedLink operation.</returns>
        public WOPIHttpResponse RevokeRestrictedLink(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiRestrictedLink)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-RestrictedLink", xwopiRestrictedLink);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.RevokeRestrictedLink));
            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.RevokeRestrictedLink);

            this.ValidateFilesCapture();

            return responseTemp;
        }

        /// <summary>
        /// This method is used to access the WOPI server's implementation of a secure store.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource URL.</param>
        /// <param name="commonHeaders">A parameter represents the common headers that contain "Authorization" header and etc.</param>
        /// <param name="xwopiApplicationId">A parameter represents the value of application ID.</param>
        /// <returns>A return value represents the http response for ReadSecureStore operation.</returns>
        public WOPIHttpResponse ReadSecureStore(string targetResourceUrl, WebHeaderCollection commonHeaders, string xwopiApplicationId)
        {
            WOPIHttpResponse responseTemp = null;

            if (null == commonHeaders)
            {
                commonHeaders = new WebHeaderCollection();
            }

            commonHeaders.Add("X-WOPI-ApplicationId", xwopiApplicationId);
            commonHeaders.Add("X-WOPI-Override", this.GetTheXWOPIOverrideHeaderValue(WOPIOperationName.ReadSecureStore));
            responseTemp = this.SendWOPIRequest(targetResourceUrl, commonHeaders, null, WOPIOperationName.ReadSecureStore);

            this.ValidateFilesCapture();
            this.ValidateReadSecureStoreResponse(responseTemp);

            return responseTemp;
        }

        #endregion

        #region protected method

        /// <summary>
        /// A method is used to get the headers value from a web exception's http response.
        /// </summary>
        /// <param name="webException">A parameter represents the web exception instance.</param>
        /// <returns>A return value represents the headers value which contain the information about the web exception.</returns>
        protected virtual string GetWebExceptionHeadersValue(WebException webException)
        {
            if (null != webException)
            {
                HttpWebResponse errorResponse = webException.Response as HttpWebResponse;
                return this.OutPutHeadersValue(errorResponse.Headers);
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// A method is used to record information about web exception which is thrown by calling protocol operations.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the request URL which cause a web exception.</param>
        /// <param name="wopiOperationName">A parameter represents the protocol operation name which sends a http request and then get a web exception.</param>
        /// <param name="webException">A parameter represents web exception instance.</param>
        protected virtual void RecordWebExceptionInformation(string targetResourceUrl, WOPIOperationName wopiOperationName, WebException webException)
        {   
            string errorHeaderValue = this.GetWebExceptionHeadersValue(webException);

            if (!string.IsNullOrEmpty(errorHeaderValue))
            {
                  this.Site.Log.Add(
                                LogEntryKind.Debug,
                                "Perform the [{0}] operation fail. Request URL:[{1}]\r\nError headers:\r\n{2}",
                                wopiOperationName,
                                targetResourceUrl,
                                errorHeaderValue);
            }
        }

        /// <summary>
        /// A method is used to send http request for MS-WOPI operation.
        /// </summary>
        /// <param name="targetResourceUrl">A parameter represents the target resource Uri.</param>
        /// <param name="headers">A parameter represents the headers which is included in the http request.</param>
        /// <param name="body">A parameter represents the body contents which is sent in the http request.</param>
        /// <param name="operationName">A parameter represents the WOPI operation which the http request belongs to.</param>
        /// <returns>A return value represents the http response of the http request which is sent with specified header, URI, body and http method. </returns>
        protected virtual WOPIHttpResponse SendWOPIRequest(string targetResourceUrl, WebHeaderCollection headers, byte[] body, WOPIOperationName operationName)
        {
            HttpWebResponse responseTemp = null;
            HttpWebRequest request = null;

            request = (HttpWebRequest)HttpWebRequest.Create(targetResourceUrl);
            request.Method = this.GetHttpMethodForWOPIOperation(operationName);
            
            // Setting the required common headers
            if (null == headers)
            {
                headers = new WebHeaderCollection();
            }

            request.Headers = headers;

            if (null == body || 0 == body.Length)
            {
                request.ContentLength = 0;
            }
            else
            {
                request.ContentLength = body.Length;
                request.ContentType = "application/binary";
                Stream stream = request.GetRequestStream();
                stream.Write(body, 0, body.Length);
            }
 
            // Get the response by default user credential
            request.Credentials = new NetworkCredential(this.defaultUserName, this.defaultPassword, this.defaultDomain);
            
            // Log the HTTP request
            this.LogHttpTransportInfo(request, body, null, operationName, false);

            try
            {
                responseTemp = request.GetResponse() as HttpWebResponse;
            }
            catch (WebException webEx)
            {
                this.Site.Log.Add(
                                LogEntryKind.Debug,
                                @"There is a WebException generated when sending WOPI request. Exception message[{0}],\r\nStackTrace:[{1}]",
                                webEx.Message,
                                webEx.StackTrace);

                this.RecordWebExceptionInformation(targetResourceUrl, operationName, webEx);
                throw;
            }

            WOPIHttpResponse wopiHttpResponse = new WOPIHttpResponse(responseTemp);

            // Log the HTTP response
            this.LogHttpTransportInfo(request, null, wopiHttpResponse, operationName);

            this.ValidateCommonMessageCapture();
           
            return wopiHttpResponse;
        }
        
        #endregion 

        #region private method

        /// <summary>
        /// A method is used to get the X-WOPI-Override header value according to the operation name.
        /// </summary>
        /// <param name="operationName">A parameter represents the operation name.</param>
        /// <returns>A return value represents the X-WOPI-Override header value which is used by specified operation.</returns>
        private string GetTheXWOPIOverrideHeaderValue(WOPIOperationName operationName)
        {
            string valueTemp = string.Empty;
            switch (operationName)
            {
                case WOPIOperationName.Lock:
                case WOPIOperationName.UnlockAndRelock:
                    {
                        valueTemp = "LOCK";
                        break;
                    }

                case WOPIOperationName.PutFile:
                    {
                        valueTemp = "PUT";
                        break;
                    }

                case WOPIOperationName.UnLock:
                    {
                        valueTemp = "UNLOCK";
                        break;
                    }

                case WOPIOperationName.PutRelativeFile:
                    {
                        valueTemp = "PUT_RELATIVE";
                        break;
                    }

                case WOPIOperationName.RefreshLock:
                    {
                        valueTemp = "REFRESH_LOCK";
                        break;
                    }

                case WOPIOperationName.ExecuteCellStorageRequest:
                case WOPIOperationName.ExecuteCellStorageRelativeRequest:
                    {
                        valueTemp = "COBALT";
                        break;
                    }

                case WOPIOperationName.DeleteFile:
                    {
                        valueTemp = "DELETE";
                        break;
                    }

                case WOPIOperationName.GetRestrictedLink:
                    {
                        valueTemp = "GET_RESTRICTED_LINK";
                        break;
                    }

                case WOPIOperationName.RevokeRestrictedLink:
                    {
                        valueTemp = "REVOKE_RESTRICTED_LINK";
                        break;
                    }

                case WOPIOperationName.ReadSecureStore:
                    {
                        valueTemp = "READ_SECURE_STORE";
                        break;
                    }

                default:
                    {
                        string errorMsg = string.Format(@"There is no valid [X-WOPI-Override] header value for operation[{0}]", operationName);
                        throw new InvalidOperationException(errorMsg);
                    }
            }

            return valueTemp;
        }

        /// <summary>
        /// A method is used to out put headers values into a string value.
        /// </summary>
        ///  <param name="headers">A parameter represents the headers' information.</param>
        /// <returns>A return value represents the string which contains the headers' name-value pairs.</returns>
        private string OutPutHeadersValue(WebHeaderCollection headers)
        {
            if (null == headers)
            {
                throw new ArgumentNullException("headers");
            }

            if (null == headers || 0 == headers.Count)
            {
                return string.Empty;
            }

            // Get each header name-value pairs from the headers' collection.
            StringBuilder strBuilder = new StringBuilder();
            foreach (string oheaderNameItem in headers.AllKeys)
            {
                string headerValueString = string.Format(
                                                     "[{0}]:({1})\r\n",
                                                     oheaderNameItem,
                                                     headers[oheaderNameItem]);

                strBuilder.Append(headerValueString);
            }

            return strBuilder.ToString();
        }

        /// <summary>
        /// A method used to log the HTTP transport information.
        /// </summary>
        /// <param name="wopiRequest">A parameter represents the request instance of a WOPI operation. It must not be null.</param>
        /// <param name="requestBody">A parameter represents the request body of a WOPI operation.</param>
        /// <param name="wopiResponse">A parameter represents the response of a WOPI operation. It must not be null if the "isResponse" parameter is true.</param>
        /// <param name="operationName">A parameter represents the WOPI operation name.</param>
        /// <param name="isResponse">A parameter represents the HTTP transport information recorded by this method whether belong to the response of a WOPI operation.</param>
        private void LogHttpTransportInfo(HttpWebRequest wopiRequest, byte[] requestBody, WOPIHttpResponse wopiResponse, WOPIOperationName operationName, bool isResponse = true)
        {
            #region validate parameters
            if (isResponse)
            {
                if (null == wopiResponse)
                {
                    throw new ArgumentNullException("wopiResponse");
                }

                // For log response, requires the request URL.
                if (null == wopiRequest)
                {
                    throw new ArgumentNullException("wopiRequest");
                }
            }
            else
            {
                if (null == wopiRequest)
                {
                    throw new ArgumentNullException("wopiRequest");
                }
            }
            #endregion 

            #region headers

            // Build headers information
            StringBuilder headerInfoBuilder = new StringBuilder();
            WebHeaderCollection headers = isResponse ? wopiResponse.Headers : wopiRequest.Headers;
            if (null != headers && 0 != headers.Count)
            {
                foreach (string oheaderNameItem in headers.AllKeys)
                {
                    string headerValueString = string.Format(
                                                         "[{0}]:({1})",
                                                         oheaderNameItem,
                                                         headers[oheaderNameItem]);

                    headerInfoBuilder.AppendLine(headerValueString);
                }
            }

            #endregion 

            #region body information

            // Build body information
            byte[] httpBodyOfResponse = null;
            if (isResponse && 0 != wopiResponse.ContentLength)
            {
                httpBodyOfResponse = WOPIResponseHelper.GetContentFromResponse(wopiResponse);
            }

            byte[] httpBody = isResponse ? httpBodyOfResponse : requestBody;
            StringBuilder bodyInfoBuilder = new StringBuilder();
            if (null != httpBody && 0 != httpBody.Length)
            {
                switch (operationName)
                {
                    case WOPIOperationName.CheckFileInfo:
                    case WOPIOperationName.ReadSecureStore:
                    case WOPIOperationName.CheckFolderInfo:
                    case WOPIOperationName.EnumerateChildren:
                        {
                            if (isResponse)
                            {
                                // log response body by JSON format
                                bodyInfoBuilder.AppendLine(Encoding.UTF8.GetString(httpBody));
                            }

                            break;
                        }

                    case WOPIOperationName.PutRelativeFile:
                        {
                            if (isResponse)
                            {
                                // log the body by JSON format
                                bodyInfoBuilder.AppendLine(Encoding.UTF8.GetString(httpBody));
                            }
                            else
                            {
                                // log the body as bytes string value
                                bodyInfoBuilder.AppendLine(WOPIResponseHelper.GetBytesStringValue(httpBody));
                            }

                            break;
                        }

                    case WOPIOperationName.GetFile:
                        {
                            if (isResponse)
                            {
                                // log the body as bytes string value
                                bodyInfoBuilder.AppendLine(WOPIResponseHelper.GetBytesStringValue(httpBody));
                            }

                            break;
                        }

                    case WOPIOperationName.PutFile:
                        {
                            if (!isResponse)
                            {
                                // log the body as bytes string value
                                bodyInfoBuilder.AppendLine(WOPIResponseHelper.GetBytesStringValue(httpBody));
                            }

                            break;
                        }
                }
            }

            #endregion 

            string credentialInfo = string.Format(
                        "User:[{0}] Domain:[{1}]",
                        this.defaultUserName,
                        this.defaultDomain);

            string logTitle = string.Format(
                                    "{0} HTTP {1}{2} for [{3}] operation:",
                                    isResponse ? "Receive" : "Sending",
                                    isResponse ? "Response" : "Request",
                                    isResponse ? string.Empty : " with " + credentialInfo,
                                    operationName);

            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine(logTitle);

            string urlInfor = string.Format("Request URL:[{0}]", wopiRequest.RequestUri.AbsoluteUri);
            logBuilder.AppendLine(urlInfor);

            string httpMethodValue = string.Format("HTTP method:[{0}]", this.GetHttpMethodForWOPIOperation(operationName));
            logBuilder.AppendLine(httpMethodValue);

            if (isResponse)
            {
                string httpStatusCodeValue = string.Format("HTTP status code:[{0}]", wopiResponse.StatusCode);
                logBuilder.AppendLine(httpStatusCodeValue);
            } 

            string headerInfo = string.Format("Headers:\r\n{0}", 0 == headerInfoBuilder.Length ? "None" : headerInfoBuilder.ToString());
            logBuilder.AppendLine(headerInfo);

            string bodyInfo = string.Format("Body:\r\n{0}", 0 == bodyInfoBuilder.Length ? "None" : bodyInfoBuilder.ToString());
            logBuilder.AppendLine(bodyInfo);

            this.Site.Log.Add(LogEntryKind.Debug, logBuilder.ToString());
        }

        /// <summary>
        /// A method used to get the HTTP method value according to the WOPI operation name.
        /// </summary>
        /// <param name="operationName">A parameter represents the WOPI operation name.</param>
        /// <returns>A return value represents the HTTP method value for the WOPI operation.</returns>
        private string GetHttpMethodForWOPIOperation(WOPIOperationName operationName)
        {
            string httpMethod = string.Empty;
            switch (operationName)
            {
                case WOPIOperationName.CheckFileInfo:
                case WOPIOperationName.CheckFolderInfo:
                case WOPIOperationName.GetFile:
                case WOPIOperationName.EnumerateChildren:
                    {
                        httpMethod = WebRequestMethods.Http.Get;
                        break;
                    }

                default:
                    {
                        httpMethod = WebRequestMethods.Http.Post;
                        break;
                    }
            }

            return httpMethod;
        }

        #endregion 
    }
}