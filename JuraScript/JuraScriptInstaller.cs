using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.IO;


namespace JuraScript
{
    [RunInstaller(true)]
    public partial class JuraScriptInstaller : System.Configuration.Install.Installer
    {
        public JuraScriptInstaller()
        {
            InitializeComponent();
        }
        public override void Install(IDictionary stateSaver)
        {
            //Add the path variable so that .jur are directly executable
            string pathext = Environment.GetEnvironmentVariable("PATHEXT");

            if (!pathext.ToUpper().Contains("JUR"))
            {
                pathext += ";.JUR";
            }
            Environment.SetEnvironmentVariable("PATHEXT", pathext, EnvironmentVariableTarget.Machine);
            base.Install(stateSaver);
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
        }
        protected override void OnAfterInstall(IDictionary savedState)
        {
            

            base.OnAfterInstall(savedState);
        }
        protected override void OnAfterUninstall(IDictionary savedState)
        {
            base.OnAfterUninstall(savedState);
        }
    }
}
