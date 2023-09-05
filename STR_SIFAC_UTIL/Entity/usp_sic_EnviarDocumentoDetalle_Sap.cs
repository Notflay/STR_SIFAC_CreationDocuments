using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_UTIL.Entity
{
    public class usp_sic_EnviarDocumentoDetalle_Sap
    {
        [MaxLength(20)]
        public string NidDoc { get; set; }

        public int OrdDet { get; set; }

        [MaxLength(18)]
        public string MatDet { get; set; }

        public decimal CanDet { get; set; }

        [MaxLength(3)]
        public string UniMed { get; set; }

        [MaxLength(4)]
        public string ClaCon { get; set; }

        public double ImpDet { get; set; }

        public double CanBas { get; set; }

        public string TexDet { get; set; }
        /* Nuevo a implementar*/
        public double DiscPrnct { get; set; }
        [Required]
        public string TaxCode { get; set; }
        public string U_BPP_OPER { get; set; }
        public int U_STR_FECodAfect { get; set; }
    }
}
