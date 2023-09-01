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
        private static bool procesoTerminado = false;
        public Service1()
        {
            InitializeComponent();
                       
        }

        protected override void OnStart(string[] args)
        {
            Global.WriteToFile("Servicio Iniciado... " + DateTime.Now);
            try
            {
                stLapso.Start();
            }
            catch (Exception ex)
            {

                this.Stop();
            }
        }

        protected override void OnStop()
        {
            stLapso.Stop();
            Global.WriteToFile("Informar detención de servicio " + DateTime.Now);
            Global.WriteToFile("Servicio detenido " + DateTime.Now);
        }

        private void IntegrarDocumentos()
        {
            try
            { 
                if (procesoTerminado) return;

                
                 procesoTerminado = true;
                 ServicioCreation servicioCreation = new ServicioCreation();
                 servicioCreation.IntegrarDocumentos().Wait();
                 servicioCreation.IntegrarEnviados().Wait();
                 servicioCreation.IntegrarRechazados().Wait();
                 servicioCreation.IntegrarCancelados().Wait();
                 procesoTerminado = false;
                
            }
            catch (Exception exec)
            {   
                Global.WriteToFile("ERROR:" + exec.Message);
                procesoTerminado = true;
            }

        }

        private void stLapso_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            IntegrarDocumentos();
        }
    }
}
