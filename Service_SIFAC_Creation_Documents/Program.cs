﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace Service_SIFAC_Creation_Documents
{
    internal static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;    
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);

            //Service1 service = new Service1();
            //service.Ejecutar();
        }
    }
}
