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
    using System.Collections.Generic;

    /// <summary>
    /// This class represents a set of PropertyRows
    /// </summary>
    public class PropertyRowSet : Node
    {
        /// <summary>
        /// The rows count
        /// </summary>
        private int count;

        /// <summary>
        /// The property rows
        /// </summary>
        private List<PropertyRow> propertyRows;

        #region Properties
        /// <summary>
        /// Gets or sets the rows count.
        /// </summary>
        public int Count
        {
            get { return this.count; }
            set { this.count = value; }
        }

        /// <summary>
        /// Gets the property rows.
        /// </summary>
        public List<PropertyRow> PropertyRows
        {
            get { return this.propertyRows; }
        }
        #endregion

        /// <summary>
        /// Parse bytes in context into a PropertyRowSetNode
        /// </summary>
        /// <param name="context">The value of Context</param>
        public override void Parse(Context context)
        {
            // No PropertyRowNode to parse
            if (this.count <= 0)
            {
                return;
            }

            // Clear PropretyRows list to store parsing result
            context.PropertyRows.Clear();

            // Parse PropertyRow one by one
            for (int i = 0; i < this.count; i++)
            {
                if (context.IsEnd())
                {
                    throw new ParseException("End prematurely");
                }

                PropertyRow propertyRow = new PropertyRow();
                propertyRow.Parse(context);
                context.PropertyRows.Add(propertyRow);
            }

            // Assign parsing result to PropertyRows
            this.propertyRows = context.PropertyRows;
        }
    }
}