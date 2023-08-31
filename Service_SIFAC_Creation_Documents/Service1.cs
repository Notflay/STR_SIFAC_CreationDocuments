using STR_SIFAC_Creation;
using STR_SIFAC_UTIL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace Service_SIFAC_Creation_Documents
{
    public partial class Service1 : ServiceBase
    {
        private System.Threading.Timer timer = null;
        private static bool procesoTerminado;
        public Service1()
        {
            InitializeComponent();
            procesoTerminado = true;
        }

        protected override void OnStart(string[] args)
        {
            Global.WriteToFile("Servicio Iniciado... " + DateTime.Now);
            try
            {

            }
            catch (Exception ex)
            {

                this.Stop();
            }
            timer = new System.Threading.Timer((g) =>
            {
                IntegrarDocumentos();
                timer.Change(Convert.ToInt32(ConfigurationManager.AppSettings["intervalo"]), Timeout.Infinite);
            }, null, 0, Timeout.Infinite);
        }

        protected override void OnStop()
        {
            Global.WriteToFile("Informar detención de servicio " + DateTime.Now);
            Global.WriteToFile("Servicio detenido " + DateTime.Now);
        }


        public void Ejecutar()
        {

            IntegrarDocumentos();

        }

        private void IntegrarDocumentos()
        {
            try
            {
                if (procesoTerminado)
                {
                    procesoTerminado = false;
                    ServicioCreation.IntegrarDocumentos().Wait();
                    ServicioCreation.IntegrarEnviados().Wait(); 
                    procesoTerminado = true;
                }
            }
            catch (Exception exec)
            {
                Global.WriteToFile("ERROR:" + exec.Message);
                procesoTerminado = true;
            }

        }
    }
}
