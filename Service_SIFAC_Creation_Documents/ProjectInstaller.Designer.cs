namespace Service_SIFAC_Creation_Documents
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.STR_SIFAC_INTEGRACION = new System.ServiceProcess.ServiceProcessInstaller();
            this.serviceinstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // STR_SIFAC_INTEGRACION
            // 
            this.STR_SIFAC_INTEGRACION.Account = System.ServiceProcess.ServiceAccount.LocalService;
            this.STR_SIFAC_INTEGRACION.Password = null;
            this.STR_SIFAC_INTEGRACION.Username = null;
            // 
            // serviceinstaller
            // 
            this.serviceinstaller.Description = "Servicio Integrador entre SAP y SIFAC";
            this.serviceinstaller.ServiceName = "Integracion Sap-Sifac";
            this.serviceinstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.serviceinstaller,
            this.STR_SIFAC_INTEGRACION});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller STR_SIFAC_INTEGRACION;
        private System.ServiceProcess.ServiceInstaller serviceinstaller;
    }
}