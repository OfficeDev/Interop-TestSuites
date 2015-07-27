//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WOPI
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// A class is used to record the logs information.
    /// </summary>
    public class LogsRecorder
    {  
        /// <summary>
        /// A dictionary instance represents the logs records identified by log record name.
        /// </summary>
        private Dictionary<string, StringBuilder> logsRecordsStorage;

        /// <summary>
        /// Initializes a new instance of the <see cref="LogsRecorder"/> class.
        /// </summary>
        public LogsRecorder()
        {
            this.logsRecordsStorage = new Dictionary<string, StringBuilder>();
        }

        /// <summary>
        /// A method is used to clean up the existed logs .
        /// </summary>
        public void CleanUpLogs()
        {
            if (this.logsRecordsStorage.Count != 0)
            {
                this.logsRecordsStorage.Clear();
            }
        }

        /// <summary>
        /// A method is used to add a log record item.
        /// </summary>
        /// <param name="logRecordName">A parameter represents the log record name, it is the unique identifier for the log record. If the left this parameter as empty, the log string will be added into the record named "DefaultLogRecord".</param>
        /// <param name="logMessage">A parameter represents the log information will be added.</param>
        public void AddLogs(string logRecordName, string logMessage)
        {
            if (string.IsNullOrEmpty(logMessage))
            {
                return;
            }

            if (string.IsNullOrEmpty(logRecordName))
            {
                logRecordName = "DefaultLogRecord";
            }

            StringBuilder logvalue = this.GetLogValueForSpecifiedLogRecord(logRecordName);

            if (null == logvalue)
            {
                logvalue = new StringBuilder();
                this.logsRecordsStorage.Add(logRecordName, logvalue);
                logvalue.AppendLine(logMessage);
            }
            else
            {
                logvalue.AppendLine(logMessage);
            }
        }

        /// <summary>
        /// A method is used to get all log record items' information.
        /// </summary>
        /// <param name="logRecordName">A parameter represents the log record name, it is the  unique identifier for the log record. If the left this parameter as empty, the log string of log record named "DefaultLogRecord" will be return.</param>
        /// <returns>A return value represents all log information of the specified log record.</returns>
        public string GetAllLogs(string logRecordName)
        {
            StringBuilder logValue = this.GetLogValueForSpecifiedLogRecord(logRecordName);
            if (null == logValue)
            {
                return string.Empty;
            }
            else
            {
                return logValue.ToString();
            }
        }

        /// <summary>
        /// A method is used to get log values for specified log record. If the expected log record does not exist for specified log record name, this method will return null.
        /// </summary>
        /// <param name="logRecordName">A parameter represents the log record name, it is the unique identifier for the log record. If the left this parameter as empty, the log string of log record named "DefaultLogRecord" will be return.</param>
        /// <returns>A return value represents the log string for specified log record. </returns>
        protected StringBuilder GetLogValueForSpecifiedLogRecord(string logRecordName)
        {
            if (string.IsNullOrEmpty(logRecordName))
            {
                logRecordName = "DefaultLogRecord";
            }

            var logValuesOflogRecord = from logRecord in this.logsRecordsStorage
                                       where logRecord.Key.Equals(logRecordName, System.StringComparison.OrdinalIgnoreCase)
                                       select logRecord.Value;

            StringBuilder logValue = null;
            if (0 == logValuesOflogRecord.Count())
            {
                if (logRecordName.Equals("DefaultLogRecord", StringComparison.OrdinalIgnoreCase))
                {
                    logValue = new StringBuilder();
                    this.logsRecordsStorage.Add(logRecordName, logValue);
                }

                return logValue;
            }
            else
            {
                logValue = logValuesOflogRecord.ElementAt<StringBuilder>(0);
            }

            return logValue;
        }
    }
}