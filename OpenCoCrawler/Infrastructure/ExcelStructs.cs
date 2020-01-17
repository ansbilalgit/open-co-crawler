using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenCoCrawler.Infrastructure
{
    public class ExcelStructs
    {
        // Structures used for conversion between data-types.
        public struct ExcelDataTypes
        {
            public const string NUMBER = "NUMBER";
            public const string DATETIME = "DATETIME";
            public const string TEXT = "TEXT"; // also works with "STRING".
        }

        public struct NETDataTypes
        {
            public const string SHORT = "int16";
            public const string INT = "int32";
            public const string LONG = "int64";
            public const string STRING = "string";
            public const string DATE = "DateTime";
            public const string BOOL = "Boolean";
            public const string DECIMAL = "decimal";
            public const string DOUBLE = "double";
            public const string FLOAT = "float";
        }
    }
}
