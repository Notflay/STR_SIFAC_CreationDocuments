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

        public static string sqlQuery { get; set; }
        public static string Almacen { get; set; }
        public static bool Pruebas { get; set; }
        public ServicioCreation()
        {
            try
            {
                // Se cambia la configuración en App.config de los parametros (Optimiza la consulta)
                OrgVen = ConfigurationManager.AppSettings["OrgVen"];
                Year = string.IsNullOrEmpty(ConfigurationManager.AppSettings["year"]) ? DateTime.UtcNow.Year.ToString() : ConfigurationManager.AppSettings["year"];
                Month = string.IsNullOrEmpty(ConfigurationManager.AppSettings["month"]) ? DateTime.UtcNow.Month.ToString() : ConfigurationManager.AppSettings["month"];
                UseSer = ConfigurationManager.AppSettings["UseSer"];
                PasSer = ConfigurationManager.AppSettings["PasSer"];
                UrlSifac = ConfigurationManager.AppSettings["urlSifac"];
                Almacen = ConfigurationManager.AppSettings["almacen"];
                Pruebas = string.IsNullOrEmpty(ConfigurationManager.AppSettings["PRUEBAS"]) ? false : ConfigurationManager.AppSettings["PRUEBAS"] == "1" ? true : false;
            }
            catch (Exception e)
            {
                WriteToFile($"Error - Servicio: " + e.Message.ToString());
            }
        }
        public async Task IntegrarDocumentos()
        {
            Connect();

            List<usp_sic_EnviarDocumento_Sap> data = NewDocuments();

            if (data.Count > 0)
            {               
                foreach (usp_sic_EnviarDocumento_Sap d in data)
                {
                    try
                    {

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
                                string serieRef = d.FolioDocRef.Remove(4);
                                string correlativoRef = d.FolioDocRef.Remove(0, 5);
                                oDocumento.UserFields.Fields.Item("U_BPP_MDTO").Value = "01";
                                oDocumento.UserFields.Fields.Item("U_BPP_MDSO").Value = serieRef; // serie 
                                oDocumento.UserFields.Fields.Item("U_BPP_MDCO").Value = correlativoRef; // correlativo

                                oDocumento.UserFields.Fields.Item("U_STR_MtvoCD").Value = d.MotDoc;

                                oDocumento.UserFields.Fields.Item("U_BPP_SDocDate").Value = QuerySql.GetDocDateRef(serieRef,correlativoRef);
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

                            oDocumento.CardCode = QuerySql.GetDocumentId(d.SolDoc);
                            oDocumento.DocDate = Convert.ToDateTime(d.FecDocFac);
                            oDocumento.TaxDate = Convert.ToDateTime(d.FecDocFac);
                            oDocumento.DocDueDate = Convert.ToDateTime(d.FecDocFac).AddDays(Convert.ToInt32(d.ConPag.Remove(0,1)));


                   
                             if (d.MonDoc != "PEN")                       
                                oDocumento.DocCurrency = "USD";
                            


                            oDocumento.Comments = d.NroPedCliente;
                            oDocumento.DocTotal = d.MonTotal;

                            int linea = 0;
                            double total = 0;
                            foreach (usp_sic_EnviarDocumentoDetalle_Sap de in d.DetDoc)
                            {
                                // Cambiar estado si ya no esta en pruebas
       
                                string matDet = Pruebas ? "VPGN00000001" : de.MatDet;
          
                                QuerySql.ValidarExistencia(matDet);
                                oItem.GetByKey(matDet);

                                if (oItem.InventoryItem == BoYesNoEnum.tYES)
                                    QuerySql.ValidarStock(matDet, Almacen, Convert.ToInt32(de.CanDet));

                                oDocumento.Lines.SetCurrentLine(linea);

                                bool esServicio = oItem.InventoryItem == BoYesNoEnum.tNO && oItem.SalesItem == BoYesNoEnum.tYES /* &&  oItem.PurchaseItem == BoYesNoEnum.tNO && oItem.GLMethod == BoGLMethods.glm_ItemLevel*/;


                                if (!esServicio)
                                    oDocumento.Lines.AccountCode = QuerySql.AccountCode();

                                oDocumento.Lines.ItemCode = matDet;
                                oDocumento.Lines.ItemDescription = de.TexDet;

                                // DATOS DE LOCALIZACION INGRESO COMO CONSTANTE
                                oDocumento.Lines.CostingCode = "0001";

                                //******** Los datos de abajo van a cambiar **************//
                                oDocumento.Lines.CostingCode2 = "400000";
                                oDocumento.Lines.CostingCode4 = "CO00CM34";
                                //*********************************************
                                oDocumento.Lines.UserFields.Fields.Item("U_TCH_N_CONT").Value = "01";


                                oDocumento.Lines.Quantity = Convert.ToDouble(de.CanDet); // Cantidad 

                                oDocumento.Lines.UnitPrice = de.TaxCode == "EXO" ? de.ImpDet / Convert.ToDouble(de.CanDet)
                                    : (de.ImpDet / 1.18) / Convert.ToDouble(de.CanDet);   // Precio Unico cantidad 
                                //oDocumento.Lines.UnitPrice = de.TaxCode == "EXO" ? de.ImpDet / Convert.ToDouble(de.CanDet)
                                //    : de.ImpDet  / Convert.ToDouble(de.CanDet);   // Precio Unico cantidad 

                                oDocumento.Lines.Price = de.TaxCode == "EXO" ? de.ImpDet : 
                                    (de.ImpDet / 1.18) / Convert.ToDouble(de.CanDet);  // Precio Unitario del producto 
                                //oDocumento.Lines.Price = de.ImpDet / Convert.ToDouble(de.CanDet);  // Precio Unitario del producto 

                                oDocumento.Lines.LineTotal = de.TaxCode == "EXO" ? de.ImpDet : de.ImpDet / 1.18; // Precio unitario del producto * cantidad 
                                //oDocumento.Lines.LineTotal = de.TaxCode == "EXO" ? de.ImpDet : de.ImpDet; // Precio unitario del producto * cantidad
                                // oDocumento.Lines.PriceAfterVAT = ''

                                if (!esServicio)
                                    oDocumento.Lines.COGSAccountCode = QuerySql.CogsAcct(Almacen);

                                oDocumento.Lines.TaxCode = de.TaxCode == "IGV" ? "IGV18" : de.TaxCode; // IGV - EXO
                                oDocumento.Lines.WarehouseCode = Almacen;
                               
                                oDocumento.Lines.DiscountPercent = de.DiscPrnct == null ? 0.0 : de.DiscPrnct;
                                oDocumento.Lines.UserFields.Fields.Item("U_BPP_OPER").Value = de.U_BPP_OPER;
                                oDocumento.Lines.UserFields.Fields.Item("U_STR_FECodAfect").Value = Convert.ToString(de.U_STR_FECodAfect);

                                total += de.ImpDet;

                                oDocumento.Lines.Add();
                                linea++;
                            }

                            string xml = oDocumento.GetAsXML();

                            if (oDocumento.Add() == 0)
                            {
                                
                                WriteToFile($"Servicio - (ObtenerDocumento): {documento } {serieDoc + "-" + correlativoDoc} " +
                                    $"creado exitosamente!");
                            }
                            else
                            {
                                WriteToFile($"Error - Servicio (ObtenerDocumento): {documento} {sboCompany.GetLastErrorDescription()}");
                            }
                        }
                        else
                        {
                            WriteToFile($"Error - Servicio (ObtenerDocumento): {documento} ya fue creado anteriormente {d.NidDoc}. Enviarlo al proveedor");
                        }
                    }
                    catch (Exception e)
                    {
                        WriteToFile($"Error - Servicio (ObtenerDocumento): {e.Message}");
                    }
                }
            }
        }       
        public async Task IntegrarEnviados()
        {

            try
            {

                oSq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} STR_Docs_Aceptados_Sifac");
                


                int recor = oSq.RecordCount;
                

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


                    while (!oSq.EoF)
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
                                ResObtener response = JsonSerializer.Deserialize<ResObtener>(cliente.SendAsync(request).Result.Content.ReadAsStringAsync().Result);
                                if (response.FlaSer)
                                {
                                    oHq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac ACE,{body["NidDoc"]},{oSq.Fields.Item(2).Value}");

                                    WriteToFile($"Servicio (ActualizarDocumento): ¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");
                                }
                                else
                                {
                                    WriteToFile($"Error - Servicio (ActualizarDocumento): {body["NidDoc"]} " + response.LogSer);
                                }
                            }

                        }
                        catch (Exception e)
                        {
                            WriteToFile("Error - Servicio (ActualizarDocumento): " + e.Message);
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

                                ResObtener response = JsonSerializer.Deserialize<ResObtener>(cliente.SendAsync(request).Result.Content.ReadAsStringAsync().Result);
                                if (response.FlaSer)
                                {
                                    oHq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac ERR,{body["NidDoc"]},{oSq.Fields.Item(3).Value}");
                                    WriteToFile($"Servicio (ActualizarDocumento): ¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");
                                }
                                else {
                                    WriteToFile($"Error - Servicio (ActualizarDocumento): {body["NidDoc"]} " + response.LogSer);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            WriteToFile("Error - Servicio (ActualizarDocumento): " + e.Message);
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

                                ResObtener response = JsonSerializer.Deserialize<ResObtener>(cliente.SendAsync(request).Result.Content.ReadAsStringAsync().Result);
                                if (response.FlaSer)
                                {
                                    oHq.DoQuery($"{(QueryPosition == 0 ? "EXEC" : "CALL")} Str_Docs_Update_Sifac BAJ,{body["NidDoc"]},{oSq.Fields.Item(3).Value}");

                                    WriteToFile($"Servicio (ActualizarDocumento): ¡Documento {body["FolDoc"]} fue actualizado a {body["StaDoc"]} exitosamente!");
                                }
                                else
                                {
                                    WriteToFile($"Error - Servicio (ActualizarDocumento): {body["NidDoc"]} " + response.LogSer);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            WriteToFile("Error - Servicio (ActualizarDocumento): " + e.Message);
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
                        WriteToFile($"Error - Servicio (ObtenerDocumento): {data.LogSer}");
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
                if (!sboCompany.Connected)
                {
                    sboCompany.Server = ConfigurationManager.AppSettings["SAP_SERVIDOR"];
                    sboCompany.CompanyDB = ConfigurationManager.AppSettings["SAP_BASE"];
                    sboCompany.DbServerType = QuerySql.GetTypeDB(ConfigurationManager.AppSettings["SAP_TIPO_BASE"]);
                    sboCompany.DbUserName = ConfigurationManager.AppSettings["SAP_DBUSUARIO"];
                    sboCompany.DbPassword = ConfigurationManager.AppSettings["SAP_DBPASSWORD"];
                    sboCompany.UserName = ConfigurationManager.AppSettings["SAP_USUARIO"];
                    sboCompany.Password = ConfigurationManager.AppSettings["SAP_PASSWORD"];
                    sboCompany.language = BoSuppLangs.ln_Spanish_La;

                    if (sboCompany.Connect() != 0)
                    {
                        WriteToFile("CONEXION-SAPConnector:" + sboCompany.GetLastErrorDescription());
                        throw new Exception(Global.sboCompany.GetLastErrorDescription());
                    }
                    else
                    {
                        WriteToFile("Conexion Exitosa");

                        oSq = sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        oHq = sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        QueryPosition = sboCompany.DbServerType == BoDataServerTypes.dst_HANADB ? 1 : 0;
                    }

                };
            }
            catch (Exception ex)
            {
                WriteToFile("Conexion :" + ex.Message);
            }

        }
     
    }
}
