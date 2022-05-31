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

            var file = vault.TryGetFileFromPath(@"C:\SOLIDWORKSPDM\Bluebyte\api\knapheide\bodies\12240859.SLDASM", out folder,40);

            file.CheckOut(folder, 0);


            Console.Read();

        }
    }
}
