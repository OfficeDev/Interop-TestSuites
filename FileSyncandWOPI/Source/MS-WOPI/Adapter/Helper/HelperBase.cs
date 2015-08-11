namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Threading;

    /// <summary>
    /// A class is used to support basic log function for other helper class implementation.
    /// </summary>
    public class HelperBase
    {   
        /// <summary>
        /// A LogsRecorder type instance which is used to record the logs information
        /// </summary>
        private static LogsRecorder logsRecorder = new LogsRecorder();

        /// <summary>
        /// A method is used to get a string which represents the logs information.
        /// </summary>
        /// <param name="helperType">A parameter represents the type information of the helper.</param>
        /// <returns>A return value represents the logs of the helper which is identified by "helperType" parameter.</returns>
        public static string GetLogs(Type helperType)
        {
            if (null != helperType)
            {
                string recordName = GetHelperRecordName(helperType);
                return logsRecorder.GetAllLogs(recordName);
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// A method is used to clean up all the logs of the helper.
        /// </summary>
        public static void CleanUpLogs()
        {
            logsRecorder.CleanUpLogs();
        }

        /// <summary>
        /// A method is used to check whether a parameter input is null or empty. If the parameter does not pass the validation, this method will throw an ArgumentNullException.
        /// </summary>
        /// <typeparam name="T">The type of the parameter which will be check.</typeparam>
        /// <param name="parameter">A parameter represents the parameter input which will be checked.</param>
        /// <param name="parameterName">A parameter represents the name of the parameter which will be check.</param>
        /// <param name="methodName">A parameter represents the name of a method whose parameter that will be checked.</param>
        public static void CheckInputParameterNullOrEmpty<T>(T parameter, string parameterName, string methodName)
        {
            if (string.IsNullOrEmpty(parameterName))
            {
                throw new ArgumentNullException("parameterName");
            }

            if (string.IsNullOrEmpty(methodName))
            {
                throw new ArgumentNullException("methodName");
            }

            Type currentParameterType = typeof(T);
            string errorMessageTemplate = @"The [{0}]type parameter[{1}] of the method [{2}] is null or empty.";
            if (null == parameter)
            {
              string errorMessage = string.Format(
                                                errorMessageTemplate,
                                                currentParameterType.FullName,
                                                parameterName,
                                                methodName);

              throw new ArgumentNullException(parameterName, errorMessage);
            }

            // If the type is string, check it whether is string.empty
            if (currentParameterType.FullName.IndexOf("System.string", StringComparison.OrdinalIgnoreCase) >= 0)
            {  
               string currentStringValue = parameter as string;
               if (string.IsNullOrEmpty(currentStringValue))
               {
                   string errorMessage = string.Format(
                                                errorMessageTemplate,
                                                currentParameterType.Name,
                                                parameterName,
                                                methodName);

                   throw new ArgumentNullException(parameterName, errorMessage);
               }
            }
        }

        /// <summary>
        /// A method is used to append logs information to the log recorder.
        /// </summary>
        /// <param name="helperType">A parameter represents the type of the current helper.</param>
        /// <param name="logMessage">A parameter represents the log information which is appended.</param>
        protected static void AppendLogs(Type helperType, string logMessage)
        {
            string helperRecordName = GetHelperRecordName(helperType);
            logsRecorder.AddLogs(helperRecordName, logMessage);
        }

        /// <summary>
        /// A method is used to append logs information to the log recorder.
        /// </summary>
        /// <param name="helperType">A parameter represents the type of the current helper.</param>
        /// <param name="timeStamp">A parameter represents the time stamp value when the logs happen.</param>
        /// <param name="logMessage">A parameter represents the log information which is appended.</param>
        /// <param name="isconvertToUTC">A parameter represents a bool value indicating the method whether convert the DateTime instance to UTC time. Default value is false, means does not convert to UTC time.</param>
        protected static void AppendLogs(Type helperType, DateTime timeStamp, string logMessage, bool isconvertToUTC = false)
        {
            string threadManagedId = Thread.CurrentThread.ManagedThreadId.ToString();
            logMessage = string.Format(
                            @"{0}({1}): CurrentThread[{2}]->{3}",
                            isconvertToUTC ? timeStamp.ToUniversalTime().ToShortTimeString() : timeStamp.ToShortTimeString(),
                            isconvertToUTC ? "UTC" : "LocalTime",
                            threadManagedId,
                            logMessage);
            AppendLogs(helperType, logMessage);
        }
 
        /// <summary>
        /// A method is used to get log record name for a helper. All log record name should be generated by this method.
        /// </summary>
        /// <param name="helperType">A parameter represents the type information of the helper.</param>
        /// <returns>A return value represents the helper record name.</returns>
        protected static string GetHelperRecordName(Type helperType)
        {
            string helperLogRecordName = string.Empty;
            helperLogRecordName = string.Format("{0}Log", helperType.Name);
            return helperLogRecordName;
        }
    }
}