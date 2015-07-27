//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCTABL
{
    /// <summary>
    /// The utility
    /// </summary>
    public class Constants
    {
        /// <summary>
        /// The field identifies that the the buffer is too small.
        /// </summary>
        public const string BufferTooSmall = "ecBufferTooSmall"; 

        /// <summary>
        /// The field identifies the folder name created to test the hierarchy table.
        /// </summary>
        public const string TestGetHierarchyTableFolderName1 = "TestGetHierarchyTableFolder1";

        /// <summary>
        /// The field identifies a sub folder name created to test the hierarchy table.
        /// </summary>
        public const string TestGetHierarchyTableFolderName2 = "TestGetHierarchyTableFolder2";

        #region Rule Name
        /// <summary>
        /// Specify the value for the rule name of OP_MARK_AS_READ rule, the default value on exchange is "MarkAsRead".
        /// </summary>
        public const string RuleNameMarkAsRead = "MarkAsRead";
        #endregion

        #region Test Data for rule data

        /// <summary>
        /// Specify the value for the PidTagRuleUserFlags, the default value on exchange is "1".
        /// </summary>
        public const string PidTagRuleUserFlags1 = "1";

        /// <summary>
        /// Specify the value for the PidTagRuleProvider, the default value on exchange is "RuleOrganizer".
        /// </summary>
        public const string PidTagRuleProvider = "RuleOrganizer";

        /// <summary>
        /// Specify the value for the PidTagRuleProviderData, the default value on exchange is "01000000010000002222222270C1E340".
        /// </summary>
        public const string PidTagRuleProviderData = "01000000010000002222222270C1E340";

        /// <summary>
        /// Specify the value for the rule condition subject, the default value on exchange is "fdx".
        /// </summary>
        public const string RuleConditionSubjectContainString = "fdx";

        /// <summary>
        /// Specify the value for ActionFlavor of Rule Action, the default value on exchange is "0".
        /// </summary>
        public const uint CommonActionFlavor = 0;

        /// <summary>
        /// Specify the value for ActionFlags of Rule Action, the default value on exchange is "0".
        /// </summary>
        public const uint RuleActionFlags = 0;
        #endregion
    }
}