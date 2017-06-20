using System;
using System.Collections.Generic;
using System.Text;

namespace FastExcel
{

    /// <summary>
    /// Exception used during loading process of defined names
    /// </summary>
    public class DefinedNameLoadException : Exception
    {
        public DefinedNameLoadException(string message, Exception innerException = null)
            : base(message, innerException)
        {

        }
    }
}
