using System;
using System.IO;
using System.IO.Compression;

namespace OutlookProfileBackup
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Chemin vers le profil Outlook de l'utilisateur courant
                string outlookProfilePath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Outlook");

                // Vérifie si le dossier Outlook existe
                if (!Directory.Exists(outlookProfilePath))
                {
                    Console.WriteLine("Le profil Outlook n'a pas été trouvé.");
                    return;
                }

                // Demander à l'utilisateur le nom du fichier ZIP
                Console.WriteLine("Entrez le nom du fichier ZIP pour la sauvegarde : ");
                string zipFileName = Console.ReadLine();

                // Vérifier et corriger le nom du fichier
                if (string.IsNullOrWhiteSpace(zipFileName))
                {
                    zipFileName = "OutlookBackup.zip";
                }
                else if (!zipFileName.EndsWith(".zip"))
                {
                    zipFileName += ".zip";
                }

                // Chemin complet vers le bureau
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string zipFilePath = Path.Combine(desktopPath, zipFileName);

                // Créer le fichier ZIP
                Console.WriteLine("Sauvegarde en cours...");
                ZipFile.CreateFromDirectory(outlookProfilePath, zipFilePath, CompressionLevel.Optimal, true);

                Console.WriteLine($"Sauvegarde terminée avec succès !");
                Console.WriteLine($"Fichier sauvegardé : {zipFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Une erreur s'est produite : {ex.Message}");
            }

            Console.WriteLine("Appuyez sur une touche pour quitter...");
            Console.ReadKey();
        }
    }
}
