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
using static STR_SIFAC_UTIL.Global;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace STR_SIFAC_Creation
{
    public class ServicioCreation
    {
        public static string OrgVen { get; set; }
        public static string Year { get; set; }
        public static string Month { get; set; }
        public static string UseSer { get; set; }
        public static string PasSer { get; set; }
        public static string UrlSifac { get; set; }

        private SqlConnection sqlConnection;
        public static string sqlQuery { get; set; }
        public ServicioCreation()
        {
            // Se cambia la configuración en App.config de los parametros (Optimiza la consulta)
            OrgVen = ConfigurationManager.AppSettings["OrgVen"].ToString();
            Year = string.IsNullOrEmpty(ConfigurationManager.AppSettings["year"]) ? DateTime.UtcNow.Year.ToString() : ConfigurationManager.AppSettings["year"];
            Month = string.IsNullOrEmpty(ConfigurationManager.AppSettings["month"]) ? DateTime.UtcNow.Month.ToString() : ConfigurationManager.AppSettings["month"];
            UseSer = ConfigurationManager.AppSettings["UseSer"].ToString();
            PasSer = ConfigurationManager.AppSettings["PasSer"].ToString();
            UrlSifac = ConfigurationManager.AppSettings["urlSifac"].ToString();

            //sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["sql"].ConnectionString);
        }
        public async Task IntegrarDocumentos()
        {

            List<usp_sic_EnviarDocumento_Sap> data = NewDocuments();

            if (data.Count > 0)
            {
                Connect();

                foreach (usp_sic_EnviarDocumento_Sap d in data)
                {

                    try
                    {
                        //string almacenOrg = "";

                        //sqlQuery = "SELECT TOP 1 WhsCode FROM OWHS";
                        //sqlConnection.Open();


                        //using (var cmd = new SqlCommand(sqlQuery, sqlConnection))
                        //{
                        //    using (var reader = cmd.ExecuteReader())
                        //    { 
                        //        if (reader.HasRows)
                        //            while (reader.Read())
                        //            {
                        //                almacenOrg = reader.GetString(0);
                        //            }
                        //    }
                        //}
                        //sqlConnection.Close();

                        oSq.DoQuery("SELECT TOP 1 WhsCode FROM OWHS");
                        string almacenOrg = oSq.Fields.Item(0).Value;

                        string tipoDoc = QuerySql.GetTipo(d.ClaDoc);
                        string serieDoc = QuerySql.GetSerie(tipoDoc);
                        string correlativoDoc = QuerySql.GetCorrelativo(serieDoc, tipoDoc);
                        string documento = tipoDoc == "01" ? "Factura" : tipoDoc == "07" ? "Nota de Credito" : "Nota de debito";
                        if (QuerySql.Validacion(d.ClaDoc, d.NidDoc))
                        {

                            Documents oDocumento = tipoDoc == "01" ? (Documents)sboCompany.GetBusinessObject(BoObjectTypes.oInvoices) :
                               tipoDoc == "07" ? (Documents)sboCompany.GetBusinessObject(BoObjectTypes.oCreditNotes) :
                               (Documents)sboCompany.GetBusinessObject(BoObjectTypes.oInvoices);

                            Items oItem = sboCompany.GetBusinessObject(BoObjectTypes.oItems);

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

                            oDocumento.DocDueDate = Convert.ToDateTime(d.FecDocFac).AddDays(Convert.ToInt32(d.ConPag.Remove(0,1)));

                            oDocumento.Comments = "Documento en estado de pruebas";

                            oDocumento.DocTotal = d.MonTotal;




                            int linea = 0;
                            double total = 0;
                            foreach (usp_sic_EnviarDocumentoDetalle_Sap de in d.DetDoc)
                            {
                                // Cambiar para produccion quitar lo comentado
                                // string matDet = string.IsNullOrEmpty(de.MatDet) ? "VPGN00000001" : de.MatDet;
                                string matDet ="VPGN00000001";
                                QuerySql.ValidarExistencia(matDet);
                                oItem.GetByKey(matDet);

                                if (oItem.InventoryItem == BoYesNoEnum.tYES)
                                    QuerySql.ValidarStock(matDet, almacenOrg, Convert.ToInt32(de.CanDet));

                                oDocumento.Lines.SetCurrentLine(linea);

                                bool esServicio = oItem.InventoryItem == BoYesNoEnum.tNO && oItem.SalesItem == BoYesNoEnum.tYES /* &&  oItem.PurchaseItem == BoYesNoEnum.tNO && oItem.GLMethod == BoGLMethods.glm_ItemLevel*/;


                                if (!esServicio)
                                    oDocumento.Lines.AccountCode = QuerySql.AccountCode();

                                oDocumento.Lines.ItemCode = string.IsNullOrEmpty(matDet) ? "VPGN00000001" : de.MatDet;
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

                                oDocumento.Lines.DiscountPercent = de.DiscPrnct == null ? 0.0 : de.DiscPrnct;
                                oDocumento.Lines.UserFields.Fields.Item("U_BPP_OPER").Value = de.U_BPP_OPER;
                                oDocumento.Lines.UserFields.Fields.Item("U_STR_FECodAfect").Value = Convert.ToString(de.U_STR_FECodAfect);

                                oDocumento.Lines.CostingCode2 = null;
                                oDocumento.Lines.CostingCode2 = null;
                                total += de.ImpDet;



                                oDocumento.Lines.Add();
                                linea++;
                            }

                            string xml = oDocumento.GetAsXML();

                            if (oDocumento.Add() == 0)
                            {
                                WriteToFile($"{documento} con correlativo {correlativoDoc} creado exitosamente!");
                            }
                            else
                            {
                                WriteToFile($"Error al crear {documento}: {sboCompany.GetLastErrorDescription()}");
                            }
                        }
                        else
                        {
                            WriteToFile($"Error: {documento} ya fue creado anteriormente {d.NidDoc}. Enviarlo al proveedor");
                        }
                    }
                    catch (Exception e)
                    {
                        WriteToFile($"Error: {e.Message}");
                    }
                }
            }
        }       
        public async Task IntegrarEnviados()
        {

            if (!sboCompany.Connected)
                Connect();
            try
            {

                oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} STR_Docs_Aceptados_Sifac");
                    

                if (oSq.RecordCount > 0)
                {


                    var body = new Dictionary<string, string>()
                    {
                            { "NidDoc", "" },
                            { "FolDoc", "" },
                            { "StaDoc", "ACE" },
                            { "UseSer", UseSer},
                            { "PasSer", PasSer}
                    };


                    while(!oSq.EoF)
                    {
                        try
                        {
                            using (var cliente = new HttpClient())
                            {
                                body["NidDoc"] = oSq.Fields.Item(0).Value;
                                body["FolDoc"] = oSq.Fields.Item(1).Value;
                                

                                var request = new HttpRequestMessage()
                                {
                                    RequestUri = new Uri(UrlSifac + "ActualizarDocumento"),
                                    Method = HttpMethod.Post,
                                    Content = new FormUrlEncodedContent(body)
                                };

                                // Actualiza el documento en SIFAC a estado ACE
                                var response = cliente.SendAsync(request).Result;
                                if (response.IsSuccessStatusCode)
                                {
                                    oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac ACE,{body["NidDoc"]},{oSq.Fields.Item(2).Value}");
                                  
                                    WriteToFile($"¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");                                                                    
                                }
                                else
                                {
                                    WriteToFile("No se pudo actualizar documento: " + response.Content.ReadAsStringAsync().Result);
                                }
                            }

                        }
                        catch (Exception e)
                        {
                            WriteToFile("Error al actualizar documento: " + e.Message);
                        }
                        finally {
                            oSq.MoveNext();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                WriteToFile(e.Message);

            }
        }
        public async Task IntegrarRechazados()
        {
            try
            {

                oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} STR_Docs_Rechazados_Sifac");

                if (oSq.RecordCount > 0)
                {
                    var body = new Dictionary<string, string>()
                    {
                            { "NidDoc", "" },
                            { "FolDoc", "" },
                            { "StaDoc", "ERR" },
                            { "TexSta",""},
                            { "UseSer", UseSer},
                            { "PasSer", PasSer}
                    };

                    while (!oSq.EoF)
                    {
                        try
                        {
                            using (var cliente = new HttpClient())
                            {
                                body["NidDoc"] = oSq.Fields.Item(0).Value;
                                body["FolDoc"] = oSq.Fields.Item(1).Value;
                                body["TexSta"] = oSq.Fields.Item(2).Value;

                                var request = new HttpRequestMessage()
                                {
                                    RequestUri = new Uri(UrlSifac + "ActualizarDocumento"),
                                    Method = HttpMethod.Post,
                                    Content = new FormUrlEncodedContent(body)
                                };

                                var response = cliente.SendAsync(request).Result;
                                if (response.IsSuccessStatusCode)
                                {
                                    oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac ERR,{body["NidDoc"]},{oSq.Fields.Item(3).Value}");
                                    WriteToFile($"¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");
                                }
                                else {
                                    WriteToFile("No se pudo actualizar documento: " + response.Content.ReadAsStringAsync().Result);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            WriteToFile("No se pudo actualizar documento: " + e.Message);
                        }
                        finally {
                            oSq.MoveNext();
                        }

                    }
                }

            }
            catch (Exception e)
            {
                WriteToFile(e.Message);
            }

        }
        public async Task IntegrarCancelados()
        {
            try
            {

                oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} STR_Docs_Cancelados_Sifac");

                if (oSq.RecordCount > 0)
                {
                    var body = new Dictionary<string, string>()
                    {
                            { "NidDoc", "" },
                            { "FolDoc", "" },
                            { "StaDoc", "BAJ" },
                            { "TexSta",""},
                            { "UseSer", UseSer},
                            { "PasSer", PasSer}
                    };

                    while (!oSq.EoF)
                    {
                        try
                        {
                            using (var cliente = new HttpClient())
                            {
                                body["NidDoc"] = oSq.Fields.Item(0).Value;
                                body["FolDoc"] = oSq.Fields.Item(1).Value;
                                body["TexSta"] = oSq.Fields.Item(2).Value;

                                var request = new HttpRequestMessage()
                                {
                                    RequestUri = new Uri(UrlSifac + "ActualizarDocumento"),
                                    Method = HttpMethod.Post,
                                    Content = new FormUrlEncodedContent(body)
                                };

                                var response = cliente.SendAsync(request).Result;
                                if (response.IsSuccessStatusCode)
                                {
                                    oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac BAJ,{body["NidDoc"]},{oSq.Fields.Item(3).Value}");

                                    WriteToFile($"¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");
                                }
                                else
                                {
                                    WriteToFile("No se pudo actualizar documento: " + response.Content.ReadAsStringAsync().Result);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            WriteToFile("No se pudo actualizar documento: " + e.Message);
                        }
                        finally
                        {
                            oSq.MoveNext();
                        }

                    }
                }

            }
            catch (Exception e)
            {
                WriteToFile(e.Message);
            }
        }
        public static List<usp_sic_EnviarDocumento_Sap> NewDocuments()
        {
            try
            {

                using (var cliente = new HttpClient())
                {
                    var body = new Dictionary<string, string>()
                    {
                        { "OrgVen", OrgVen},
                        { "Ejercicio", Convert.ToInt32(Year).ToString()},
                        { "Periodo", Convert.ToInt32(Month).ToString()},
                        { "StaDoc", "PEN"},
                        { "UseSer", UseSer},
                        { "PasSer", PasSer}

                    };

                    var request = new HttpRequestMessage()
                    {
                        RequestUri = new Uri(UrlSifac + "ObtenerDocumento"),
                        Method = HttpMethod.Post,
                        Content = new FormUrlEncodedContent(body)
                    };

                    var response = cliente.SendAsync(request).Result;
                    ResObtener data = JsonSerializer.Deserialize<ResObtener>(response.Content.ReadAsStringAsync().Result);

                    if (data.FlaSer)
                        return data.DatSer;
                    else
                        WriteToFile($"ERROR: {data.LogSer}");
                    throw new Exception();
                }
            }
            catch (Exception e)
            {
                WriteToFile(e.Message.ToString());
                throw new Exception();
            }

        }
        public static void Connect()
        {
            try
            {
                sboCompany.Server = ConfigurationManager.AppSettings["SAP_SERVIDOR"];
                sboCompany.CompanyDB = ConfigurationManager.AppSettings["SAP_BASE"];
                sboCompany.DbServerType = QuerySql.GetTypeDB(ConfigurationManager.AppSettings["SAP_TIPO_BASE"]);
                sboCompany.DbUserName = ConfigurationManager.AppSettings["SAP_DBUSUARIO"];
                sboCompany.DbPassword = ConfigurationManager.AppSettings["SAP_DBPASSWORD"];
                sboCompany.UserName = ConfigurationManager.AppSettings["SAP_USUARIO"];
                sboCompany.Password = ConfigurationManager.AppSettings["SAP_PASSWORD"];
                sboCompany.language = BoSuppLangs.ln_Spanish_La;

                Conexion();
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message);
            }

        }
        public static void Conexion()
        {
            try
            {
                if (!sboCompany.Connected)
                {
                    if (sboCompany.Connect() != 0)
                    {
                        WriteToFile("CONEXION-SAPConnector:" + sboCompany.GetLastErrorDescription());
                        throw new Exception(Global.sboCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        WriteToFile("CONEXION EXITOSA");

                        oSq = sboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        QueryPosition = sboCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB ? 1 : 0;
                    }
                }
                

            }
            catch (Exception ex)
            {

                WriteToFile("CONEXION :" + ex.Message);
            }
        }


    }
}
