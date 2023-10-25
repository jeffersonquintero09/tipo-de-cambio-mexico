using System;
using System.Globalization;
using System.IO;
using System.Net;
using Newtonsoft.Json.Linq;
using SAPbobsCOM;

namespace CrearSN
{
    class Program
    {
        public static SAPbobsCOM.Company myCompany = null;
        static void Main(string[] args)
        {
            ActualizarTipoDeCambio();
        }

        public static void ActualizarTipoDeCambio()
        {
            try
            {
                SAPbobsCOM.SBObob oSBObob;
                SAPbobsCOM.Recordset oRecordSet;
                if (ConexionSAP())
                {
                    string url = "https://sidofqa.segob.gob.mx/dof/sidof/indicadores/";
                    string jsonResult;

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Method = "GET";

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        using (Stream stream = response.GetResponseStream())
                        {
                            StreamReader reader = new StreamReader(stream);
                            jsonResult = reader.ReadToEnd();
                        }
                    }

                    // Analiza la respuesta JSON para obtener el valor del tipo de cambio correspondiente al código 34542
                    JObject data = JObject.Parse(jsonResult);
                    JArray listaIndicadores = (JArray)data["ListaIndicadores"];
                    double tipoCambio = 0;

                    foreach (var indicador in listaIndicadores)
                    {
                        if ((int)indicador["codTipoIndicador"] == 158)
                        {
                            // Obtener el valor del tipo de cambio como cadena
                            string tipoCambioStr = (string)indicador["valor"];

                            // Convertir el valor del tipo de cambio a double
                            if (double.TryParse(tipoCambioStr, NumberStyles.Float, CultureInfo.GetCultureInfo("es-MX"), out double tipoCambioDouble))
                            {
                                tipoCambio = tipoCambioDouble;
                            }
                            else
                            {
                                Console.WriteLine("No se pudo convertir el valor del tipo de cambio a double: " + tipoCambioStr);
                            }

                            break;
                        }
                    }


                    Console.WriteLine("Tipo de cambio: " + tipoCambio);

                    if (myCompany.Connected)
                    {
                        oSBObob = myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        oRecordSet = myCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet = oSBObob.GetLocalCurrency();
                        oRecordSet = oSBObob.GetSystemCurrency();

                        // Formatear el valor del tipo de cambio manualmente
                        string tipoCambioFormatted = tipoCambio.ToString("0.0000", CultureInfo.InvariantCulture);

                        oSBObob.SetCurrencyRate("USD", DateTime.Now, tipoCambio, true);

                        if (myCompany.Connected == true)
                        {
                            myCompany.Disconnect();
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Environment.Exit(0);
            }
            
        }

        public static bool ConexionSAP()
        {
            bool respuesta = false;
            try
            {
                myCompany = new SAPbobsCOM.Company();
                myCompany.Server = ""; //IP o Nombre del dominio del servidor de base de datos
                myCompany.DbServerType = BoDataServerTypes.dst_MSSQL2017; //Tipo de base de datos
                // Todas las conexiones que necesitemos               
                myCompany.CompanyDB = ""; //Nombre base de datos SAP
                myCompany.UserName = ""; //Nombre usuario SAP
                myCompany.Password = ""; //Contraseña usuario SAP
                myCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La;

                int iRet = myCompany.Connect();

                if (iRet == 0)
                {
                    Console.WriteLine("Conexión exitosa a SAP");
                    respuesta = true;
                }
                else
                {
                    Console.WriteLine(myCompany.GetLastErrorDescription().ToString());
                }

                return respuesta;

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());

                return respuesta;
            }
        }
    }
}
