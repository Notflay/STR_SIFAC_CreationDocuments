using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_UTIL.Entity
{
    public class usp_sic_EnviarDocumentoCuota_Sap
    {
        public string NidDoc { get; set; }
        public int NroCuota { get; set; }
        public string FecPagoCuota { get; set; }
        public double ImpDet { get; set; }
    }
}
