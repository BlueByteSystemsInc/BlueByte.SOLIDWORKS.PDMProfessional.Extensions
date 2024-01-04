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
        private static EdmVault5 vault;
        private static string originalFileLocation;
        private static string vaultFile;
        private static int handle = int.MinValue;

        static void Main(string[] args)
        {

            handle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle.ToInt32();

            vault = new EdmVault5();

            vault.LoginAuto("bluebyte", handle);


            originalFileLocation = @"C:\Users\jlili\Desktop\P017-525-0033Default.xlsx";
            vaultFile = @"C:\SOLIDWORKSPDM\Bluebyte\API\pdm2excel\thumnails example\P017-525-0033Default.xlsx";


            TestAddFileNonExistingPDM();
            TestAddFileNonExistingPDM();
            TestAddFileWhileExistingCheckedIn();
            TestAddFileWhileExistingCheckedOut();
            Console.Read();

        }


        public static void TestAddFileNonExistingPDM()
        {
            IEdmFolder5 originalFolder;
            var file = vault.TryGetFileFromPath(vaultFile, out originalFolder);

            if (file != null)
                originalFolder.DeleteFile(handle, file.ID, true);

            if (originalFolder != null)
            originalFolder.Refresh();

            string error;

            vault.AddFile(originalFileLocation, vaultFile, handle, out error);

            string assert = string.IsNullOrEmpty(error) ? "Pass" : $"Fail {error}";

            Console.WriteLine($"{System.Reflection.MethodBase.GetCurrentMethod().Name}: {assert}");

        }

        public static void TestAddFileWhileExistingCheckedIn()
        {
            IEdmFolder5 originalFolder;
            var file = vault.TryGetFileFromPath(vaultFile, out originalFolder);

            if (file == null)
                Console.WriteLine($"{System.Reflection.MethodBase.GetCurrentMethod().Name}: Test method failed because it could not find file.");
            
            string error;

            vault.AddFile(originalFileLocation, vaultFile, handle, out error);

            string assert = string.IsNullOrEmpty(error) ? "Pass" : $"Fail {error}";

            Console.WriteLine($"{System.Reflection.MethodBase.GetCurrentMethod().Name}: {assert}");

        }

        public static void TestAddFileWhileExistingCheckedOut()
        {
            IEdmFolder5 originalFolder;
            var file = vault.TryGetFileFromPath(vaultFile, out originalFolder);

            if (file == null)
                Console.WriteLine($"{System.Reflection.MethodBase.GetCurrentMethod().Name}: Test method failed because it could not find file.");

            if (file.IsLocked == false)
                file.LockFile(originalFolder.ID, handle);

            string error;

            vault.AddFile(originalFileLocation, vaultFile, handle, out error);

            string assert = string.IsNullOrEmpty(error) ? "Pass" : $"Fail {error}";

            Console.WriteLine($"{System.Reflection.MethodBase.GetCurrentMethod().Name}: {assert}");

        }


    }
}
