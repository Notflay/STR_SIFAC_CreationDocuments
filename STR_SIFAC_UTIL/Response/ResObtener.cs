using STR_SIFAC_UTIL.Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_UTIL.Response
{
    public class ResObtener
    {
        public bool FlaSer { get; set; }
        public string LogSer { get; set; }

        public List<usp_sic_EnviarDocumento_Sap> DatSer { get; set; }

    }
}
