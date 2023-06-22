using BlueByte.SOLIDWORKS.PDMProfessional.Extensions;
using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox
{
    class Program
    {
        private static IEdmFolder5 folder;

        static void Main(string[] args)
        {

            var vault = new EdmVault5();

            vault.LoginAuto("bluebyte", 0);

            var file = vault.TryGetFileFromPath(@"C:\SOLIDWORKSPDM\Bluebyte\api\sandbox\Grill Assembly\Support_Frame_End_&.SLDASM", out folder,40);

            file.Transition(folder.ID, "Major Engineering Change","Dispatch Target 001", "Sandbox transition", 0);

            file.CheckOut(folder, 0);


            Console.Read();

        }
    }
}
