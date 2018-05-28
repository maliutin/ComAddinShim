using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using NetOffice.Tools;
using NetOffice.WordApi.Tools;

namespace ISI.CdmNextAddin.WordAddin
{
    [COMAddin("Word Addin for CDMNext", "This Addin provides access to CEIC Data Services", 3)]
    [CustomUI("RibbonUI.xml", true), ComVisible(true), RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [Guid("0C32BB29-BE62-4db8-B700-85EF29B56830"), ProgId("_CDMNext.WordAddin")]
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
