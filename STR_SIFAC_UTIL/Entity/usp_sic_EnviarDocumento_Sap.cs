using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_UTIL.Entity
{
    public class usp_sic_EnviarDocumento_Sap
    {
        [MaxLength(20)]
        public string NidDoc { get; set; }

        [MaxLength(4)]
        public string ClaDoc { get; set; }

        [MaxLength(4)]
        public string OrgVen { get; set; }

        [MaxLength(4)]
        public string CanDic { get; set; }

        [MaxLength(10)]
        public string RefDoc { get; set; }

        [MaxLength(11)]
        public string SolDoc { get; set; }

        [MaxLength(10)]
        public string FecDocFac { get; set; }

        [MaxLength(4)]
        public string ConPag { get; set; }

        [MaxLength(3)]
        public string MotDoc { get; set; }

        [MaxLength(5)]
        public string MonDoc { get; set; }

        [MaxLength(35)]
        public string NroPedCliente { get; set; }

        public string TexRef { get; set; }

        [MaxLength(10)]
        public string StaDoc { get; set; }

        [MaxLength(10)]
        public string DocSap { get; set; }

        [MaxLength(10)]
        public string NumDoc { get; set; }

        [MaxLength(13)]
        public string FolioDoc { get; set; }

        [MaxLength(2)]
        public string Sector { get; set; }

        [MaxLength(10)]
        public string DestMerc { get; set; }

        public string UsuarioSiFac { get; set; }

        public double Impuesto { get; set; }

        public double MonTotal { get; set; }

        public List<usp_sic_EnviarDocumentoDetalle_Sap> DetDoc { get; set; }
        /* NUEVO */
        public string ForPago { get; set; }

        public List<usp_sic_EnviarDocumentoCuota_Sap> CuoDoc { get; set; }
    }
}
