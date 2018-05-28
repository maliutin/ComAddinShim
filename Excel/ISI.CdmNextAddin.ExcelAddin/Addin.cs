using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using NetOffice.Tools;
using NetOffice.ExcelApi.Tools;

namespace ISI.CdmNextAddin.ExcelAddin
{
    [COMAddin("Excel Addin for CDMNext", "This Addin provides access to CEIC Data Services", 3)]
    [CustomUI("RibbonUI.xml", true), RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [Guid("7F85420B-ADC1-4398-92F6-CBA2C24D5280"), ProgId("_CDMNext.ExcelAddin")]
    public class Addin : COMAddin
    {
        public Addin()
        {
            this.OnConnection += new OnConnectionEventHandler(Addin_OnConnection);
        }

        void Addin_OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            WriteText("Addin_OnConnection");
            WriteText(Environment.Version.ToString());
        }

        private void WriteText(String text)
        {
            String desctopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            String filePath = System.IO.Path.Combine(desctopDir, "001.txt");

            System.IO.File.AppendAllText(filePath, text + Environment.NewLine);
        }
    }
}
