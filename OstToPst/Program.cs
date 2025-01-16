using System;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices.Marshal;

namespace OSTtoPSTConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OST to PST Converter");
            Console.WriteLine("-------------------");

            try
            {
                Console.Write("Enter OST file path: ");
                string ostPath = Console.ReadLine();

                Console.Write("Enter destination PST path: ");
                string pstPath = Console.ReadLine();

                ConvertOSTtoPST(ostPath, pstPath);
                
                Console.WriteLine("Conversion completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        static void ConvertOSTtoPST(string ostPath, string pstPath)
        {
            Application outlook = null;
            Namespace mapi = null;

            try
            {
                // Initialize Outlook
                outlook = new Application();
                mapi = outlook.GetNamespace("MAPI");

                // Create new PST
                mapi.AddStore(pstPath);
                Store pstStore = mapi.Stores[pstPath];

                // Open OST
                Store ostStore = mapi.AddStore(ostPath);
                Folder ostRoot = ostStore.GetRootFolder();

                // Copy folders recursively
                CopyFolders(ostRoot, pstStore.GetRootFolder());

                // Cleanup
                mapi.RemoveStore(ostRoot);
            }
            finally
            {
                // Release COM objects
                if (mapi != null) Marshal.ReleaseComObject(mapi);
                if (outlook != null) Marshal.ReleaseComObject(outlook);
            }
        }

        static void CopyFolders(Folder sourceFolder, Folder destFolder)
        {
            // Copy all items in current folder
            foreach (object item in sourceFolder.Items)
            {
                item.Copy().Move(destFolder);
                Marshal.ReleaseComObject(item);
            }

            // Process subfolders
            foreach (Folder subFolder in sourceFolder.Folders)
            {
                Folder newFolder = destFolder.Folders.Add(subFolder.Name);
                CopyFolders(subFolder, newFolder);
                Marshal.ReleaseComObject(subFolder);
            }
        }
    }
}