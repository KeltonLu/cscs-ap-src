using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace CIMSService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {
            //if checked the service is stopped then start it
            ServiceController mySC = new ServiceController(this.serviceInstaller1.ServiceName);
            if (mySC.Status.Equals(ServiceControllerStatus.Stopped))
            {
                mySC.Start();
            }
        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }

        private void serviceInstaller1_BeforeUninstall(object sender, InstallEventArgs e)
        {
            //if checked the service is stopped then start it
            ServiceController mySC = new ServiceController(this.serviceInstaller1.ServiceName);
            if (mySC.Status.Equals(ServiceControllerStatus.Running))
            {
                mySC.Stop();
            }
        }
    }
}