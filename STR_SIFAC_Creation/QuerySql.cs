using SAPbobsCOM;
using STR_SIFAC_UTIL;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace STR_SIFAC_Creation
{
    public class QuerySql
    {

        public QuerySql()
        {

        }

        public static BoDataServerTypes GetTypeDB(string db)
        {
            try
            {
                switch (db)
                {
                    case "HANA":
                        return BoDataServerTypes.dst_HANADB;
                    case "SQL14":
                        return BoDataServerTypes.dst_MSSQL2014;
                    case "SQL16":
                        return BoDataServerTypes.dst_MSSQL2016;
                    case "SQL17":
                        return BoDataServerTypes.dst_MSSQL2017;
                    default:
                        throw new ArgumentException($"Tipo de base de datos no válido {db}. Intentar con(HANA,SQL14,SQL16,SQL17)");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string GetDocumentId(string doc)
        {
            try
            {
                if (doc.Length < 10)
                    return $"E{doc}";
                return $"C{doc}";    
            }
            catch (Exception e)
            {

                throw new Exception(e.Message.ToString());
            }  
        }


        public static void ValidarExistencia(string code)
        {
            try
            {
                Items item = Global.sboCompany.GetBusinessObject(BoObjectTypes.oItems);

                if (!item.GetByKey(code))
                    throw new Exception("El articulo " + code + " no está registrado en esta sociedad");

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string GetUserTable(SAPbobsCOM.UserTable userTable, string itemCof)
        {
            try
            {
                if (userTable != null)
                {
                    if (userTable.GetByKey(itemCof))
                    {
                        return userTable.UserFields.Fields.Item("U_STR_FEValor").Value;
                    }
                }
                return string.Empty;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ValidarStock(string code, string idAlmacen, int quantity)
        {
            try
            {
                Recordset rs = Global.sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string codigoAlmacen = GetAlmacen(idAlmacen);
                string query = "SELECT \"OnHand\" FROM OITW as T0 INNER JOIN OWHS AS T1 ON T0.\"WhsCode\" = T1.\"WhsCode\" WHERE \"ItemCode\" = '" + code + "' and T1.\"WhsCode\" = '" + codigoAlmacen + "' ";

                rs.DoQuery(query);

                if (rs.RecordCount > 0)
                {
                    double cantidadSAP = rs.Fields.Item(0).Value;
                    if (cantidadSAP < quantity)
                        throw new Exception("El artículo " + code + " no tiene stock suficiente en el almacén " + codigoAlmacen);
                }
                else
                    throw new Exception("El artículo " + code + " no está asociado al almacen " + codigoAlmacen);
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static string GetAlmacen(string id)
        {
            try
            {
                Recordset oRs = Global.sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = string.Format("select \"WhsCode\" FROM OWHS WHERE \"WhsCode\" = '{0}' ", id);
                oRs.DoQuery(query);

                if (oRs.RecordCount > 0)
                    return oRs.Fields.Item(0).Value;
                else
                    throw new Exception("El almacen " + id + " de BSale no está asociado a ningún almacen en SAP B1");
            }
            catch (Exception)
            {
                throw;
            }
        }



        public static string GetTipo(string tipo)
        {
            try
            {
                switch (tipo)
                {
                    case "ZPVA":
                        return "01";
                    case "ZSNC":
                        return "07";
                    case "ZSND":
                        return "08";
                    default:
                        throw new Exception("No se encontro el tipo de documento");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static string GetSerie(string tipo)
        {
            try
            {
                Global.oSq.DoQuery($"SELECT TOP 1 \"U_BPP_NDSD\" FROM \"@BPP_NUMDOC\" WHERE \"U_BPP_NDTD\" = '{tipo}' AND LEFT(\"U_BPP_NDSD\",1) = 'F' ORDER BY CAST(\"Code\" AS INT) DESC");
                if (Global.oSq.RecordCount > 0)
                    return Global.oSq.Fields.Item(0).Value;
                else
                    throw new Exception("No se encontro niguna serie, registrar en la tabla @BPP_NUMDOC");
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static string GetAcctCode()
        {
            try
            {
                Global.oSq.DoQuery("SELECT TOP 1 \"AcctCode\" FROM INV1 ORDER BY \"DocDate\" DESC");
                return Global.oSq.Fields.Item(0).Value;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static string AccountCode()
        {
            try
            {
                Recordset oRs = Global.sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = "SELECT TOP 1 \"AcctCode\" FROM INV1 ORDER BY \"DocDate\" DESC";
                oRs.DoQuery(query);

                return oRs.Fields.Item(0).Value;

            }
            catch (Exception)
            {

                throw;
            }
        }
        public static bool Validacion(string tipo, string niDoc)
        {

            try
            {
                string table = tipo == "ZSNC" ? "ORIN" : "OINV";
           
                Global.oSq.DoQuery($"SELECT \"DocEntry\" FROM {table} WHERE \"U_STR_NidDoc\" = '{niDoc}'");
                if (Global.oSq.RecordCount > 0)
                    return false;
                return true;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static string GetCorrelativo(string Serie, string tipoDoc)
        {
            try
            {
                string table = tipoDoc == "07" ? "ORIN" : "OINV";
                Global.oSq.DoQuery($"SELECT MAX(CAST(\"U_BPP_MDCD\"AS INT)) FROM {table} WHERE \"U_BPP_MDSD\" = '{Serie}'");
                int correlativo = Global.oSq.RecordCount == 0 ? 1 : Convert.ToInt32(Global.oSq.Fields.Item(0).Value) + 1;

                return Convert.ToString(correlativo);

            }
            catch (Exception)
            {
                throw new Exception("No se encontró un correlativo con la serie");
                throw;
            }
        }




        public static string CogsAcct(string idalmacen)
        {
            try
            {
                Recordset oRs = Global.sboCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = "SELECT TOP 1 \"BalInvntAc\" FROM OWHS WHERE \"WhsCode\" = '" + idalmacen + "' ";
                oRs.DoQuery(query);

                return oRs.Fields.Item(0).Value;

            }
            catch (Exception)
            {

                throw;
            }
        }

    }
}
