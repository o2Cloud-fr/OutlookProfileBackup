using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OutlookProfileBackup
{
    class Program
    {
        static int _totalFiles  = 0;
        static int _totalErrors = 0;
        static readonly string Timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.Title = "Outlook Profile Backup - Automatic";

            PrintHeader();
            Log(ConsoleColor.Cyan, "INFO", "Demarrage de la sauvegarde automatique...");
            Log(ConsoleColor.Cyan, "INFO", $"Date    : {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            Log(ConsoleColor.Cyan, "INFO", $"Machine : {Environment.MachineName}");
            Log(ConsoleColor.Cyan, "INFO", $"Utilisateur : {Environment.UserName}");
            Console.WriteLine();

            string zipFilePath = null;
            string exportedPst = null;

            try
            {
                // ── ETAPE 1 : Fermer Outlook ──────────────────────────────────────────
                Step(1, 5, "Fermeture d'Outlook");
                if (IsOutlookRunning())
                {
                    Log(ConsoleColor.Yellow, "INFO", "Outlook est ouvert, fermeture en cours...");
                    CloseOutlook();
                    System.Threading.Thread.Sleep(4000);
                    if (IsOutlookRunning())
                        throw new Exception("Outlook n'a pas pu etre ferme. Fermez-le manuellement et relancez.");
                }
                Log(ConsoleColor.Green, "OK", "Outlook est ferme.");

                // ── ETAPE 2 : Exporter le PST via COM ────────────────────────────────
                Step(2, 5, "Export PST depuis Outlook (COM)");
                exportedPst = ExportPstViaInterop();

                if (!string.IsNullOrEmpty(exportedPst) && File.Exists(exportedPst))
                    Log(ConsoleColor.Green, "OK", $"PST exporte ({FormatSize(new FileInfo(exportedPst).Length)}) : {exportedPst}");
                else
                    Log(ConsoleColor.Yellow, "WARN", "Export PST echoue ou aucune boite trouvee. Continuation...");

                // Fermer Outlook apres l'export Interop et attendre que le PST soit deverrouille
                if (IsOutlookRunning())
                {
                    Log(ConsoleColor.Cyan, "INFO", "Fermeture d'Outlook apres export...");
                    CloseOutlook();
                    System.Threading.Thread.Sleep(4000);
                }

                // Attendre que le fichier PST soit accessible (max 30 secondes)
                if (!string.IsNullOrEmpty(exportedPst) && File.Exists(exportedPst))
                {
                    Log(ConsoleColor.Cyan, "INFO", "Attente du deverrouillage du PST...");
                    int waited = 0;
                    while (waited < 30000)
                    {
                        try
                        {
                            using (var fs = File.Open(exportedPst, FileMode.Open, FileAccess.Read, FileShare.None))
                                break;
                        }
                        catch
                        {
                            System.Threading.Thread.Sleep(1000);
                            waited += 1000;
                        }
                    }
                    Log(ConsoleColor.Green, "OK", "PST deverrouille et pret a etre compresse.");
                }

                // ── ETAPE 3 : Preparer le fichier ZIP ────────────────────────────────
                Step(3, 5, "Preparation du fichier ZIP");
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string zipName     = $"OutlookBackup_{Environment.UserName}_{Timestamp}.zip";
                zipFilePath        = Path.Combine(desktopPath, zipName);
                Log(ConsoleColor.Cyan, "INFO", $"Destination : {zipFilePath}");

                // ── ETAPE 4 : Compression ─────────────────────────────────────────────
                Step(4, 5, "Compression de tous les elements");

                using (var zip = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
                {
                    Console.WriteLine();
                    _totalFiles += BackupPstFiles(zip, exportedPst);
                    _totalFiles += BackupSignatures(zip);
                    _totalFiles += BackupTemplates(zip);
                    _totalFiles += BackupStationery(zip);
                    _totalFiles += BackupRegistry(zip);
                    _totalFiles += BackupRules(zip);
                    WriteBackupManifest(zip, zipFilePath, exportedPst);
                }

                // Supprimer le PST temporaire
                if (!string.IsNullOrEmpty(exportedPst) && File.Exists(exportedPst)
                    && exportedPst.Contains("_TMP_"))
                {
                    try { File.Delete(exportedPst); Log(ConsoleColor.Cyan, "INFO", "PST temporaire supprime."); }
                    catch { }
                }

                // ── ETAPE 5 : Rapport final ───────────────────────────────────────────
                Step(5, 5, "Rapport final");
                Console.WriteLine();
                PrintSeparator();
                Log(ConsoleColor.Green,  "SUCCES", "Sauvegarde Outlook terminee avec succes !");
                PrintSeparator();
                Console.WriteLine();
                Console.WriteLine($"  Fichier ZIP    : {zipFilePath}");
                Console.WriteLine($"  Taille ZIP     : {FormatSize(new FileInfo(zipFilePath).Length)}");
                Console.WriteLine($"  Fichiers       : {_totalFiles}");
                Console.WriteLine($"  Avertissements : {_totalErrors}");
                Console.WriteLine($"  Duree          : {DateTime.Now:HH:mm:ss}");
                Console.WriteLine();
                Console.WriteLine("  CONTENU DU ZIP :");
                Console.WriteLine("  ├─ PST/           -> Donnees email Outlook (.pst)");
                Console.WriteLine("  ├─ Signatures/    -> Signatures email");
                Console.WriteLine("  ├─ Templates/     -> Modeles Outlook (.oft)");
                Console.WriteLine("  ├─ Stationery/    -> Papier a lettre Outlook");
                Console.WriteLine("  ├─ Registry/      -> Profils et comptes (registre)");
                Console.WriteLine("  ├─ Rules/         -> Regles de messagerie (.rwz)");
                Console.WriteLine("  └─ RESTORE.txt    -> Guide de restauration complet");
                Console.WriteLine();
                Console.WriteLine("  RESTAURATION RAPIDE :");
                Console.WriteLine("  1. Double-cliquez Registry\\OutlookProfiles.reg");
                Console.WriteLine("  2. Copiez PST\\ -> %LocalAppData%\\Microsoft\\Outlook\\");
                Console.WriteLine("  3. Outlook -> Fichier -> Ouvrir -> .pst");
                Console.WriteLine("  4. Copiez Signatures\\ -> %AppData%\\Microsoft\\Signatures\\");
                PrintSeparator();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                PrintSeparator();
                Log(ConsoleColor.Red, "ERREUR FATALE", ex.Message);
                PrintSeparator();

                // Nettoyage en cas d'erreur
                if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
                    try { File.Delete(zipFilePath); } catch { }
                if (!string.IsNullOrEmpty(exportedPst) && File.Exists(exportedPst))
                    try { File.Delete(exportedPst); } catch { }
            }

            Console.WriteLine();
            Console.WriteLine("  Fermeture dans 10 secondes...");
            System.Threading.Thread.Sleep(10000);
        }

        // ─────────────────────────────────────────────────────────────────────────
        // EXPORT PST VIA COM LATE-BINDING (pas de reference Interop necessaire)
        // ─────────────────────────────────────────────────────────────────────────
        static string ExportPstViaInterop()
        {
            object outlookApp = null;
            object nameSpace  = null;

            string pstPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft", "Outlook",
                $"OutlookBackup_{Timestamp}_TMP_.pst");

            try
            {
                Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                {
                    Log(ConsoleColor.Red, "ERREUR", "Outlook non installe ou introuvable via COM.");
                    return null;
                }

                Log(ConsoleColor.Cyan, "INFO", "Connexion a Outlook via COM...");
                outlookApp = Activator.CreateInstance(outlookType);
                nameSpace  = outlookApp.GetType().InvokeMember(
                    "GetNamespace", BindingFlags.InvokeMethod, null, outlookApp,
                    new object[] { "MAPI" });

                nameSpace.GetType().InvokeMember(
                    "Logon", BindingFlags.InvokeMethod, null, nameSpace,
                    new object[] { Type.Missing, Type.Missing, false, false });

                Log(ConsoleColor.Cyan, "INFO", "Creation du fichier PST de destination...");
                nameSpace.GetType().InvokeMember(
                    "AddStore", BindingFlags.InvokeMethod, null, nameSpace,
                    new object[] { pstPath });

                object pstRootFolder = GetPstRootFolderByPath(nameSpace, pstPath);
                if (pstRootFolder == null)
                    throw new Exception("Impossible de localiser le PST cree.");

                object stores     = nameSpace.GetType().InvokeMember("Stores", BindingFlags.GetProperty, null, nameSpace, null);
                int    storeCount = (int)stores.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, stores, null);

                Log(ConsoleColor.Cyan, "INFO", $"{storeCount} store(s) trouve(s). Copie en cours...");

                for (int i = 1; i <= storeCount; i++)
                {
                    object s  = stores.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, stores, new object[] { i });
                    string fp = s.GetType().InvokeMember("FilePath", BindingFlags.GetProperty, null, s, null)?.ToString() ?? "";

                    if (fp.Equals(pstPath, StringComparison.OrdinalIgnoreCase))
                    {
                        Marshal.ReleaseComObject(s);
                        continue;
                    }

                    string name = s.GetType().InvokeMember("DisplayName", BindingFlags.GetProperty, null, s, null)?.ToString() ?? "Boite";
                    Log(ConsoleColor.Cyan, "COPY", $"{name}");

                    try
                    {
                        object root = s.GetType().InvokeMember("GetRootFolder", BindingFlags.InvokeMethod, null, s, null);
                        root.GetType().InvokeMember("CopyTo", BindingFlags.InvokeMethod, null, root, new object[] { pstRootFolder });
                        Marshal.ReleaseComObject(root);
                        Log(ConsoleColor.Green, "OK", $"{name} copie.");
                    }
                    catch (Exception ex)
                    {
                        Log(ConsoleColor.Yellow, "WARN", $"{name} ignore : {ex.Message}");
                        _totalErrors++;
                    }

                    Marshal.ReleaseComObject(s);
                }

                Marshal.ReleaseComObject(stores);
                Marshal.ReleaseComObject(pstRootFolder);

                // Detacher le PST avant de le zipper
                object rootForRemove = GetPstRootFolderByPath(nameSpace, pstPath);
                if (rootForRemove != null)
                {
                    nameSpace.GetType().InvokeMember("RemoveStore", BindingFlags.InvokeMethod, null, nameSpace, new object[] { rootForRemove });
                    Marshal.ReleaseComObject(rootForRemove);
                }

                return pstPath;
            }
            catch (Exception ex)
            {
                Log(ConsoleColor.Red, "ERREUR", $"Export PST echoue : {ex.Message}");
                _totalErrors++;
                return null;
            }
            finally
            {
                if (nameSpace  != null) try { Marshal.ReleaseComObject(nameSpace);  } catch { }
                if (outlookApp != null) try { Marshal.ReleaseComObject(outlookApp); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        static object GetPstRootFolderByPath(object nameSpace, string pstPath)
        {
            object stores = nameSpace.GetType().InvokeMember("Stores", BindingFlags.GetProperty, null, nameSpace, null);
            int count = (int)stores.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, stores, null);
            for (int i = 1; i <= count; i++)
            {
                object s  = stores.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, stores, new object[] { i });
                string fp = s.GetType().InvokeMember("FilePath", BindingFlags.GetProperty, null, s, null)?.ToString() ?? "";
                if (fp.Equals(pstPath, StringComparison.OrdinalIgnoreCase))
                {
                    object root = s.GetType().InvokeMember("GetRootFolder", BindingFlags.InvokeMethod, null, s, null);
                    Marshal.ReleaseComObject(s);
                    Marshal.ReleaseComObject(stores);
                    return root;
                }
                Marshal.ReleaseComObject(s);
            }
            Marshal.ReleaseComObject(stores);
            return null;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // SAUVEGARDE PST
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupPstFiles(ZipArchive zip, string exportedPst = null)
        {
            Log(ConsoleColor.Cyan, "INFO", "Recherche des fichiers PST...");

            var pstFiles = new List<string>();

            // Ajouter EN PREMIER le PST exporte par l'outil (priorite absolue)
            if (!string.IsNullOrEmpty(exportedPst) && File.Exists(exportedPst))
            {
                pstFiles.Add(exportedPst);
                Log(ConsoleColor.Green, "INFO", $"PST exporte detecte : {Path.GetFileName(exportedPst)}");
            }

            // Rechercher aussi les PST existants sur le disque
            var searchPaths = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook"),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            foreach (string sp in searchPaths)
                if (Directory.Exists(sp))
                    pstFiles.AddRange(SafeGetFiles(sp, "*.pst"));

            pstFiles = pstFiles.Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            if (pstFiles.Count == 0)
            {
                Log(ConsoleColor.Yellow, "WARN", "Aucun fichier PST trouve.");
                _totalErrors++;
                return 0;
            }

            int count = 0;
            foreach (string pst in pstFiles)
            {
                try
                {
                    long pstSize = new FileInfo(pst).Length;
                    Log(ConsoleColor.Cyan, "INFO", $"Ajout du PST au ZIP : {Path.GetFileName(pst)} ({FormatSize(pstSize)})...");

                    // Lire le PST en memoire et l'ecrire manuellement dans le ZIP
                    // (evite les problemes de compression/verrouillage sur fichiers binaires)
                    var entry = zip.CreateEntry("PST/" + Path.GetFileName(pst), CompressionLevel.NoCompression);
                    entry.LastWriteTime = File.GetLastWriteTime(pst);

                    using (var entryStream = entry.Open())
                    using (var fileStream = new FileStream(pst, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        byte[] buffer = new byte[65536]; // 64 Ko par chunk
                        int bytesRead;
                        long totalCopied = 0;
                        while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            entryStream.Write(buffer, 0, bytesRead);
                            totalCopied += bytesRead;
                        }
                        Log(ConsoleColor.Green, "PST", $"{Path.GetFileName(pst)}  ({FormatSize(totalCopied)} copies dans le ZIP)");
                    }
                    count++;
                }
                catch (Exception ex)
                {
                    Log(ConsoleColor.Red, "ERREUR", $"PST ignore ({Path.GetFileName(pst)}) : {ex.Message}");
                    _totalErrors++;
                }
            }
            return count;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // SAUVEGARDE SIGNATURES
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupSignatures(ZipArchive zip)
        {
            Log(ConsoleColor.Cyan, "INFO", "Sauvegarde des signatures...");
            string sigPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Signatures");
            if (!Directory.Exists(sigPath))
            {
                Log(ConsoleColor.Yellow, "WARN", "Dossier Signatures introuvable.");
                return 0;
            }
            int count = 0;
            foreach (string file in Directory.GetFiles(sigPath, "*", SearchOption.AllDirectories))
            {
                try
                {
                    zip.CreateEntryFromFile(file, "Signatures/" + file.Substring(sigPath.Length).TrimStart('\\', '/'), CompressionLevel.Optimal);
                    Log(ConsoleColor.Green, "SIG", Path.GetFileName(file));
                    count++;
                }
                catch { _totalErrors++; }
            }
            return count;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // SAUVEGARDE TEMPLATES
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupTemplates(ZipArchive zip)
        {
            Log(ConsoleColor.Cyan, "INFO", "Sauvegarde des modeles (.oft)...");
            string templatePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates");
            if (!Directory.Exists(templatePath)) return 0;
            int count = 0;
            foreach (string file in Directory.GetFiles(templatePath, "*.oft", SearchOption.AllDirectories))
            {
                try
                {
                    zip.CreateEntryFromFile(file, "Templates/" + Path.GetFileName(file), CompressionLevel.Optimal);
                    Log(ConsoleColor.Green, "OFT", Path.GetFileName(file));
                    count++;
                }
                catch { _totalErrors++; }
            }
            return count;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // SAUVEGARDE PAPIER A LETTRE (Stationery)
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupStationery(ZipArchive zip)
        {
            Log(ConsoleColor.Cyan, "INFO", "Sauvegarde du papier a lettre...");
            string stationeryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Stationery");
            if (!Directory.Exists(stationeryPath)) return 0;
            int count = 0;
            foreach (string file in Directory.GetFiles(stationeryPath, "*", SearchOption.AllDirectories))
            {
                try
                {
                    zip.CreateEntryFromFile(file, "Stationery/" + Path.GetFileName(file), CompressionLevel.Optimal);
                    Log(ConsoleColor.Green, "STN", Path.GetFileName(file));
                    count++;
                }
                catch { _totalErrors++; }
            }
            return count;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // SAUVEGARDE REGLES (.rwz)
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupRules(ZipArchive zip)
        {
            Log(ConsoleColor.Cyan, "INFO", "Sauvegarde des regles de messagerie (.rwz)...");
            string outlookPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook");
            if (!Directory.Exists(outlookPath)) return 0;
            int count = 0;
            foreach (string file in SafeGetFiles(outlookPath, "*.rwz"))
            {
                try
                {
                    zip.CreateEntryFromFile(file, "Rules/" + Path.GetFileName(file), CompressionLevel.Optimal);
                    Log(ConsoleColor.Green, "RWZ", Path.GetFileName(file));
                    count++;
                }
                catch { _totalErrors++; }
            }
            return count;
        }

        // ─────────────────────────────────────────────────────────────────────────
        // EXPORT REGISTRE
        // ─────────────────────────────────────────────────────────────────────────
        static int BackupRegistry(ZipArchive zip)
        {
            Log(ConsoleColor.Cyan, "INFO", "Export du registre Outlook...");
            string officeVersion = null;
            foreach (string v in new[] { "16.0", "15.0", "14.0", "17.0" })
            {
                using (var key = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{v}\Outlook\Profiles"))
                    if (key != null) { officeVersion = v; break; }
            }

            if (officeVersion == null)
            {
                Log(ConsoleColor.Yellow, "WARN", "Cle de registre Outlook non trouvee.");
                _totalErrors++;
                return 0;
            }

            string tempRegFile = Path.Combine(Path.GetTempPath(), $"OutlookProfiles_{Timestamp}.reg");
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "reg.exe",
                    Arguments = $"export \"HKCU\\Software\\Microsoft\\Office\\{officeVersion}\\Outlook\\Profiles\" \"{tempRegFile}\" /y",
                    UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true
                };
                var proc = Process.Start(psi);
                proc.WaitForExit(10000);

                if (proc.ExitCode == 0 && File.Exists(tempRegFile))
                {
                    zip.CreateEntryFromFile(tempRegFile, "Registry/OutlookProfiles.reg", CompressionLevel.Optimal);
                    Log(ConsoleColor.Green, "REG", $"OutlookProfiles.reg (Office {officeVersion}) - {FormatSize(new FileInfo(tempRegFile).Length)}");
                    File.Delete(tempRegFile);
                    return 1;
                }
                Log(ConsoleColor.Yellow, "WARN", $"Export registre echoue (code {proc.ExitCode}).");
                _totalErrors++;
                return 0;
            }
            catch (Exception ex)
            {
                Log(ConsoleColor.Red, "ERREUR", $"Registre ignore : {ex.Message}");
                _totalErrors++;
                if (File.Exists(tempRegFile)) try { File.Delete(tempRegFile); } catch { }
                return 0;
            }
        }

        // ─────────────────────────────────────────────────────────────────────────
        // FICHIER RESTORE.txt inclus dans le ZIP
        // ─────────────────────────────────────────────────────────────────────────
        static void WriteBackupManifest(ZipArchive zip, string zipFilePath, string pstPath)
        {
            string content = $@"====================================================
  GUIDE DE RESTAURATION OUTLOOK - {DateTime.Now:dd/MM/yyyy HH:mm}
====================================================

INFORMATIONS DE SAUVEGARDE
  Machine     : {Environment.MachineName}
  Utilisateur : {Environment.UserName}
  Date        : {DateTime.Now:dd/MM/yyyy HH:mm:ss}
  Fichiers    : {_totalFiles}

----------------------------------------------------
ETAPE 1 — PROFILS ET COMPTES (Registry)
----------------------------------------------------
1. Ouvrez le dossier Registry\ dans ce ZIP
2. Double-cliquez sur OutlookProfiles.reg
3. Confirmez l'import dans le registre Windows
=> Vos comptes email et profils Outlook sont restaures

----------------------------------------------------
ETAPE 2 — FICHIERS DE DONNEES (.pst)
----------------------------------------------------
1. Copiez les fichiers du dossier PST\ vers :
   %LocalAppData%\Microsoft\Outlook\
2. Ouvrez Outlook
3. Fichier -> Ouvrir et exporter -> Ouvrir un fichier
   de donnees Outlook (.pst)
4. Selectionnez le .pst copie
=> Vos emails, contacts et calendriers sont restaures

----------------------------------------------------
ETAPE 3 — SIGNATURES EMAIL (Signatures)
----------------------------------------------------
1. Copiez tout le contenu du dossier Signatures\ vers :
   %AppData%\Microsoft\Signatures\
=> Vos signatures sont restaurees

----------------------------------------------------
ETAPE 4 — MODELES OUTLOOK (Templates)
----------------------------------------------------
1. Copiez tout le contenu du dossier Templates\ vers :
   %AppData%\Microsoft\Templates\
=> Vos modeles .oft sont restaures

----------------------------------------------------
ETAPE 5 — REGLES DE MESSAGERIE (Rules)
----------------------------------------------------
1. Copiez les fichiers .rwz du dossier Rules\ vers :
   %LocalAppData%\Microsoft\Outlook\
2. Dans Outlook : Fichier -> Gerer les regles et alertes
   -> Options -> Importer les regles
=> Vos regles de tri automatique sont restaurees

====================================================
  Sauvegarde realisee par OutlookProfileBackup
  github.com/o2Cloud-fr/OutlookProfileBackup
====================================================
";
            var entry = zip.CreateEntry("RESTORE.txt", CompressionLevel.Optimal);
            using (var writer = new StreamWriter(entry.Open(), System.Text.Encoding.UTF8))
                writer.Write(content);

            Log(ConsoleColor.Green, "TXT", "RESTORE.txt genere dans le ZIP.");
        }

        // ─────────────────────────────────────────────────────────────────────────
        // HELPERS
        // ─────────────────────────────────────────────────────────────────────────
        static bool IsOutlookRunning() =>
            Process.GetProcessesByName("OUTLOOK").Length > 0;

        static void CloseOutlook()
        {
            foreach (var proc in Process.GetProcessesByName("OUTLOOK"))
            {
                try { proc.CloseMainWindow(); proc.WaitForExit(5000); if (!proc.HasExited) proc.Kill(); }
                catch { }
            }
        }

        static IEnumerable<string> SafeGetFiles(string rootPath, string pattern)
        {
            IEnumerable<string> files = Enumerable.Empty<string>();
            try { files = Directory.GetFiles(rootPath, pattern); } catch { yield break; }
            foreach (var f in files) yield return f;

            IEnumerable<string> subDirs = Enumerable.Empty<string>();
            try { subDirs = Directory.GetDirectories(rootPath); } catch { yield break; }
            foreach (string dir in subDirs)
            {
                try { if ((File.GetAttributes(dir) & FileAttributes.ReparsePoint) != 0) continue; } catch { continue; }
                foreach (var f in SafeGetFiles(dir, pattern)) yield return f;
            }
        }

        static string FormatSize(long bytes)
        {
            if (bytes >= 1_073_741_824) return $"{bytes / 1_073_741_824.0:F1} Go";
            if (bytes >= 1_048_576)     return $"{bytes / 1_048_576.0:F1} Mo";
            if (bytes >= 1_024)         return $"{bytes / 1_024.0:F0} Ko";
            return $"{bytes} o";
        }

        static void Log(ConsoleColor color, string tag, string message)
        {
            Console.Write("  [");
            Console.ForegroundColor = color;
            Console.Write($"{tag,-8}");
            Console.ResetColor();
            Console.WriteLine($"] {message}");
        }

        static void Step(int current, int total, string name)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"  ══ ETAPE {current}/{total} : {name.ToUpper()} ══");
            Console.ResetColor();
        }

        static void PrintHeader()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("  ╔══════════════════════════════════════════╗");
            Console.WriteLine("  ║      OUTLOOK PROFILE BACKUP TOOL   v2    ║");
            Console.WriteLine("  ║           Sauvegarde Automatique         ║");
            Console.WriteLine("  ║         github.com/o2Cloud-fr            ║");
            Console.WriteLine("  ╚══════════════════════════════════════════╝");
            Console.ResetColor();
            Console.WriteLine();
        }

        static void PrintSeparator()
        {
            Console.WriteLine("  ──────────────────────────────────────────");
        }
    }
}
