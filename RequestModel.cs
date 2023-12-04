using NPOI.Util;
using Org.BouncyCastle.Asn1.Ocsp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcel
{
    public class RequestModel
    {
        public byte[] AttendanceWSFile { get; set; }

        public string Unit { get; set; }
        public string ModuleType { get ; set ; }

        public string RequestId { get ; set ; }
        public DateTime AttendanceDate { get; set; }
    }
}
