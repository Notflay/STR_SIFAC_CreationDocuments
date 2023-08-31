using SAPbobsCOM;
using STR_SIFAC_UTIL;
using STR_SIFAC_UTIL.Entity;
using STR_SIFAC_UTIL.Response;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;


namespace STR_SIFAC_Creation
{
    public class ServicioCreation
    {

        public async static Task IntegrarDocumentos()
        {

            List<usp_sic_EnviarDocumento_Sap> data = NewDocuments();

            if (data.Count > 0)
            {
                Connect();
                Conexion();

                foreach (usp_sic_EnviarDocumento_Sap d in data)
                {

                    try
                    {                      
                        Global.oSq.DoQuery("SELECT TOP 1 \"WhsCode\" FROM \"OWHS\"");

                        string almacenOrg = Global.oSq.Fields.Item(0).Value;

                        string tipoDoc = QuerySql.GetTipo(d.ClaDoc);
                        string serieDoc = QuerySql.GetSerie(tipoDoc);
                        string correlativoDoc = QuerySql.GetCorrelativo(serieDoc, tipoDoc);
                        string documento = tipoDoc == "01" ? "Factura" : tipoDoc == "07" ? "Nota de Credito" : "Nota de debito";
                        if (QuerySql.Validacion(d.ClaDoc, d.NidDoc))
                        {

                            Documents oDocumento = tipoDoc == "01" ? (Documents)Global.sboCompany.GetBusinessObject(BoObjectTypes.oInvoices) :
                               tipoDoc == "07" ? (Documents)Global.sboCompany.GetBusinessObject(BoObjectTypes.oCreditNotes) :
                               (Documents)Global.sboCompany.GetBusinessObject(BoObjectTypes.oInvoices);

                            Items oItem = Global.sboCompany.GetBusinessObject(BoObjectTypes.oItems);

                            oDocumento.DocumentSubType = tipoDoc == "08" ? BoDocumentSubType.bod_DebitMemo : BoDocumentSubType.bod_None;

                            oDocumento.UserFields.Fields.Item("U_BPP_MDTD").Value = tipoDoc;
                            oDocumento.UserFields.Fields.Item("U_BPP_MDSD").Value = serieDoc;
                            oDocumento.UserFields.Fields.Item("U_BPP_MDCD").Value = correlativoDoc;
                            oDocumento.UserFields.Fields.Item("U_STR_NidDoc").Value = d.NidDoc;

                            oDocumento.UserFields.Fields.Item("U_STR_Sifac_Estado").Value = d.StaDoc;

                            oDocumento.DocType = BoDocumentTypes.dDocument_Items;

                            oDocumento.DocObjectCode = tipoDoc == "01" ? BoObjectTypes.oInvoices
                                : tipoDoc == "07" ? BoObjectTypes.oCreditNotes : BoObjectTypes.oInvoices;

                            oDocumento.HandWritten = BoYesNoEnum.tNO;
                            oDocumento.UserFields.Fields.Item("U_STR_FE_FormaPago").Value = string.IsNullOrEmpty(d.ForPago) ? "1" : d.ForPago == "1" ? "1" : "2";

                            if (tipoDoc == "08" || tipoDoc == "07")
                            {
                                oDocumento.UserFields.Fields.Item("U_BPP_MDTO").Value = "01";
                                oDocumento.UserFields.Fields.Item("U_BPP_MDSO").Value = d.FolioDoc.Remove(4);
                                oDocumento.UserFields.Fields.Item("U_BPP_MDCO").Value = d.FolioDoc.Remove(0, 5);

                                oDocumento.UserFields.Fields.Item("U_STR_MtvoCD").Value = $"0{d.MotDoc.Remove(0, 2)}";


                            };

                            if (oDocumento.UserFields.Fields.Item("U_STR_FE_FormaPago").Value == "2" && tipoDoc == "01")
                            {
                                Document_Installments oIns;
                                oDocumento.ApplyTaxOnFirstInstallment = BoYesNoEnum.tYES;

                                int lineaCu = 0;
                                foreach (usp_sic_EnviarDocumentoCuota_Sap cuota_Sap in d.CuoDoc)
                                {
                                    oIns = oDocumento.Installments;
                                    oDocumento.Installments.SetCurrentLine(lineaCu);
                                    oDocumento.Installments.DueDate = Convert.ToDateTime(cuota_Sap.FecPagoCuota);
                                    //oDocumento.Installments.Percentage = Math.Round((cuota_Sap.ImpDet/d.MonTotal)*100,2);
                                    oDocumento.Installments.Total = cuota_Sap.ImpDet;

                                    if (lineaCu != 1)
                                        oDocumento.Installments.Add();
                                    lineaCu++;
                                }
                            }

                            oDocumento.CardCode = $"C{d.SolDoc}";

                            oDocumento.DocDate = Convert.ToDateTime(d.FecDocFac);
                            oDocumento.TaxDate = Convert.ToDateTime(d.FecDocFac);

                            oDocumento.DocDueDate = Convert.ToDateTime(d.FecDocFac).AddDays(Convert.ToInt32(d.ConPag));

                            oDocumento.Comments = "Documento en estado de pruebas";

                            oDocumento.DocTotal = d.MonTotal;


                            

                            int linea = 0;
                            double total = 0;
                            foreach (usp_sic_EnviarDocumentoDetalle_Sap de in d.DetDoc)
                            {
                                QuerySql.ValidarExistencia(de.MatDet);
                                oItem.GetByKey(de.MatDet);

                                if (oItem.InventoryItem == BoYesNoEnum.tYES)
                                    QuerySql.ValidarStock(de.MatDet, almacenOrg, Convert.ToInt32(de.CanDet));

                                oDocumento.Lines.SetCurrentLine(linea);

                                bool esServicio = oItem.InventoryItem == BoYesNoEnum.tNO && oItem.SalesItem == BoYesNoEnum.tYES /* &&  oItem.PurchaseItem == BoYesNoEnum.tNO && oItem.GLMethod == BoGLMethods.glm_ItemLevel*/;


                                if (!esServicio)
                                    oDocumento.Lines.AccountCode = QuerySql.AccountCode();

                                oDocumento.Lines.ItemCode = de.MatDet;
                                oDocumento.Lines.Quantity = Convert.ToDouble(de.CanDet); // Cantidad 
                                oDocumento.Lines.UnitPrice = de.ImpDet / Convert.ToDouble(de.CanDet);   // Precio Unico cantidad / preciototal
                                oDocumento.Lines.Price = de.ImpDet;

                                if (!esServicio)
                                    oDocumento.Lines.COGSAccountCode = QuerySql.CogsAcct(almacenOrg);

                                oDocumento.Lines.TaxCode = de.TaxCode; // IGV - EXO
                                                                       // oDocumento.Lines.WarehouseCode = almacenOrg;
                                oDocumento.Lines.LineTotal = de.ImpDet;
                                oDocumento.Lines.CostingCode = null;
                                oDocumento.Lines.CostingCode2 = null;

                                oDocumento.Lines.DiscountPercent = de.DiscPrnct;
                                oDocumento.Lines.UserFields.Fields.Item("U_BPP_OPER").Value = de.U_BPP_OPER;
                                oDocumento.Lines.UserFields.Fields.Item("U_STR_FECodAfect").Value = Convert.ToString(de.U_STR_FECodAfect);

                                oDocumento.Lines.CostingCode2 = null;
                                oDocumento.Lines.CostingCode2 = null;
                                total += de.ImpDet;



                                //oDocumento.Lines.SerialNumbers.SetCurrentLine(0);
                                //oDocumento.Lines.SerialNumbers.BaseLineNumber = linea;
                                //oDocumento.Lines.SerialNumbers.ManufacturerSerialNumber = "1";

                                //oDocumento.Lines.SerialNumbers.SystemSerialNumber = 3;
                                //oDocumento.Lines.SerialNumbers.ManufacturerSerialNumber = "11200999";
                                //oDocumento.Lines.SerialNumbers.InternalSerialNumber = "11200999";
                                //oDocumento.Lines.SerialNumbers.BaseLineNumber = 1;
                                //oDocumento.Lines.SerialNumbers.ManufacturerSerialNumber = "10900837";
                                //oDocumento.Lines.SerialNumbers.InternalSerialNumber = "11100899";
                                //oDocumento.Lines.SerialNumbers.InternalSerialNumber = "11200999";
                                //oDocumento.Lines.SerialNumbers.Add();

                                oDocumento.Lines.Add();
                                linea++;
                            }



                            string xml = oDocumento.GetAsXML();



                            if (oDocumento.Add() == 0)
                            {
                                Global.WriteToFile($"{documento} con correlativo {correlativoDoc} creado exitosamente!");
                            }
                            else
                            {
                                Global.WriteToFile($"Error al crear {documento}: {Global.sboCompany.GetLastErrorDescription()}");
                            }
                        }
                        else
                        {
                            Global.WriteToFile($"Error: {documento} ya fue creado anteriormente {d.NidDoc}. Enviarlo al proveedor");
                        }
                    }
                    catch (Exception e)
                    {
                        Global.WriteToFile($"Error: {e.Message}");
                    }
                }
            }
        }

        public static List<usp_sic_EnviarDocumento_Sap> NewDocuments()
        {
            try
            {
                // Se cambia la configuración en App.config de los parametros (Optimiza la consulta)

                string OrgVen = ConfigurationManager.AppSettings["OrgVen"].ToString();
                string year = string.IsNullOrEmpty(ConfigurationManager.AppSettings["year"]) ? DateTime.UtcNow.Year.ToString() : ConfigurationManager.AppSettings["year"];
                string month = string.IsNullOrEmpty(ConfigurationManager.AppSettings["month"]) ? DateTime.UtcNow.Month.ToString() : ConfigurationManager.AppSettings["month"];
                string UseSer = ConfigurationManager.AppSettings["UseSer"].ToString();
                string PasSer = ConfigurationManager.AppSettings["PasSer"].ToString();
                string urlSifac = ConfigurationManager.AppSettings["urlSifac"].ToString();

                using (var cliente = new HttpClient())
                {
                    var body = new Dictionary<string, string>()
                    {
                        { "OrgVen",OrgVen},
                        { "Ejercicio", Convert.ToInt32(year).ToString()},
                        { "Periodo", Convert.ToInt32(month).ToString()},
                        { "StaDoc", "PEN"},
                        { "UseSer", UseSer},
                        { "PasSer", PasSer}

                    };

                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(urlSifac + "obtener"),
                        Method = HttpMethod.Post,
                        Content = new FormUrlEncodedContent(body)
                    };

                    var response = cliente.SendAsync(request).Result;
                    ResObtener data = JsonSerializer.Deserialize<ResObtener>(response.Content.ReadAsStringAsync().Result);

                    if (data.FlaSer)
                        return data.DatSer;
                    else
                        Global.WriteToFile("ERROR: FLASER FALSE");
                    throw new Exception();
                }
            }
            catch (Exception e)
            {
                Global.WriteToFile(e.Message.ToString());
                throw new Exception();
            }

        }

        public async static Task IntegrarEnviados()
        {
            Global.oSq.DoQuery("SELECT \"U_STR_NidDoc\",\"U_BPP_MDSD\" + '-' + \"U_bPP_MDCD\" AS \"FolioDoc\" FROM OINV WHERE \"U_STR_Estado\" = 'E' AND \"U_STR_Sifac_Estado\" = 'PEN' \r\nAND ISNULL(\"U_STR_CdgHash\",'X') <> 'X' AND DATEDIFF(DAY, DocDate, GETDATE()) < 3 UNION ALL \r\nSELECT \"U_STR_NidDoc\",\"U_BPP_MDSD\" + '-'+ \"U_bPP_MDCD\" AS \"FolioDoc\" FROM ORIN WHERE \"U_STR_Estado\" = 'E' AND \"U_STR_SIfac_Estado\" = 'PEN' \r\nAND ISNULL(\"U_STR_CdgHash\",'X') <> 'X' AND DATEDIFF(DAY, DocDate, GETDATE()) < 3 ");

            while (Global.oSq.EoF)
            {
                dynamic a = Global.oSq.Fields.Item(0).Value;
            }
        
        }
       

        public static void Connect()
        {
            try
            {

                Global.sboCompany.Server = ConfigurationManager.AppSettings["SAP_SERVIDOR"];
                Global.sboCompany.CompanyDB = ConfigurationManager.AppSettings["SAP_BASE"];
                Global.sboCompany.DbServerType = ConfigurationManager.AppSettings["SAP_TIPO_BASE"] == "HANA" ? SAPbobsCOM.BoDataServerTypes.dst_HANADB : SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
                Global.sboCompany.DbUserName = ConfigurationManager.AppSettings["SAP_DBUSUARIO"];
                Global.sboCompany.DbPassword = ConfigurationManager.AppSettings["SAP_DBPASSWORD"];
                Global.sboCompany.UserName = ConfigurationManager.AppSettings["SAP_USUARIO"];
                Global.sboCompany.Password = ConfigurationManager.AppSettings["SAP_PASSWORD"];
                Global.sboCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La;

            }
            catch (Exception ex)
            {
                string mensaje = ex.Message;
                throw;
            }
        }


        public static void Conexion()
        {
            try
            {

                string a = Global.sboCompany.Server;
                string a1 = Global.sboCompany.CompanyDB;
                string a2 = Global.sboCompany.DbUserName;
                string a3 = Global.sboCompany.DbPassword;
                string a4 = Global.sboCompany.UserName;
                string a5 = Global.sboCompany.Password;
                if (Global.sboCompany.Connect() != 0)
                {
                    Global.WriteToFile("CONEXION-SAPConnector:" + Global.sboCompany.GetLastErrorDescription());
                    throw new Exception(Global.sboCompany.GetLastErrorDescription());
                }
                else
                {
                    Global.WriteToFile("CONEXION EXITOSA");

                    Global.oSq = Global.sboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Global.QueryPosition = Global.sboCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? 1 : 0;
                }

            }
            catch (Exception ex)
            {

                Global.WriteToFile("CONEXION :" + ex.Message);
            }
        }
    }
}
