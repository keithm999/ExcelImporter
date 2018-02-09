using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Exceptions
{
    [Serializable]
    public class WorksheetNotFoundException : Exception
    {

        public WorksheetNotFoundException() { }

        public WorksheetNotFoundException(string message) : base(message) { }

        public WorksheetNotFoundException(string message, Exception innerException) : base(message, innerException) { }

        protected WorksheetNotFoundException(SerializationInfo info, StreamingContext ctxt) : base(info, ctxt) { }


    }
}
