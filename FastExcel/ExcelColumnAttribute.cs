using System;

namespace FastExcel
{
    /// <summary>
    /// Add a custom name to the field
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Column name in  Excel
        /// </summary>
        public string Name { get; set; }
    }
}