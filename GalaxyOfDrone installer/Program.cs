using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO.Compression;
using IWshRuntimeLibrary; // Vergeet niet deze namespace toe te voegen voor snelkoppelingen

namespace WoWInstaller
{
    class Program
    {
        private static string downloadUrl = "https://wheeoo.org/wp-content/uploads/3.3.5a.zip";
        private static string downloadPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WoWInstaller.zip");
        private static string extractPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WoW");
        private const long MaxFileSizeInBytes = 26L * 1024 * 1024 * 1024; // 26 GB in bytes

        static async Task Main(string[] args)
        {
            // Console instellingen
            Console.Title = "World of Warcraft Installer";
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Clear();

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Welcome to the World of Warcraft Installer!\n");

            // Controleer of het bestand bestaat en of de grootte minder is dan 26 GB
            if (System.IO.File.Exists(downloadPath))
            {
                var fileInfo = new FileInfo(downloadPath);
                if (fileInfo.Length >= MaxFileSizeInBytes)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("File already exists and is large enough, skipping download.\n");
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("File exists but is smaller than 26 GB, downloading again...\n");
                    System.IO.File.Delete(downloadPath); // Verwijder het bestaande bestand als het kleiner is dan 26 GB
                    await DownloadFileAsync(downloadUrl, downloadPath); // Download opnieuw
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("File not found, downloading...\n");
                await DownloadFileAsync(downloadUrl, downloadPath); // Download het bestand
            }

            // Maak de map aan waar we de bestanden uit willen pakken
            if (!Directory.Exists(extractPath))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Creating extraction folder...\n");
                Directory.CreateDirectory(extractPath);
            }

            // Pak het bestand uit
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\nExtracting ZIP file...\n");
            if (ExtractZip(downloadPath, extractPath))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("ZIP file extracted successfully!\n");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Failed to extract ZIP file.\n");
                return;
            }

            // Probeer WoW.exe te starten
            string wowExePath = Path.Combine(extractPath, "World of Warcraft 3.3.5a", "WoW.exe");
            if (System.IO.File.Exists(wowExePath))  // Gebruik volledig gekwalificeerde naam
            {
                // Maak snelkoppeling
                CreateShortcut(wowExePath, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WoW.lnk"));

                // Start het spel
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\nStarting World of Warcraft...\n");
                System.Diagnostics.Process.Start(wowExePath);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("World of Warcraft started successfully!\n");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("WoW.exe not found in the extracted files.\n");
            }
        }

        private static async Task DownloadFileAsync(string url, string path)
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
                    {
                        response.EnsureSuccessStatusCode();
                        long totalBytes = response.Content.Headers.ContentLength ?? 0;
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.WriteLine($"Downloading {totalBytes / (1024 * 1024)} MB...\n");

                        using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            long totalDownloaded = 0;
                            byte[] buffer = new byte[8192]; // Buffer van 8KB

                            using (var stream = await response.Content.ReadAsStreamAsync())
                            {
                                int bytesRead;
                                while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                                {
                                    await fileStream.WriteAsync(buffer, 0, bytesRead);
                                    totalDownloaded += bytesRead;

                                    // Print de voortgang van de download
                                    Console.SetCursorPosition(0, Console.CursorTop);
                                    Console.Write($"Downloading: {totalDownloaded / (1024 * 1024)} MB of {totalBytes / (1024 * 1024)} MB...");
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("\nDownload completed!\n");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error downloading the file: {ex.Message}\n");
            }
        }

        private static bool ExtractZip(string zipPath, string extractTo)
        {
            try
            {
                // Pak het zipbestand uit naar de opgegeven map
                ZipFile.ExtractToDirectory(zipPath, extractTo);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting ZIP file: {ex.Message}\n");
                return false;
            }
        }

        private static void CreateShortcut(string targetFile, string shortcutLocation)
        {
            try
            {
                // Maak een snelkoppeling
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Creating shortcut on the desktop...\n");

                WshShell wshShell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)wshShell.CreateShortcut(shortcutLocation);
                shortcut.TargetPath = targetFile;
                shortcut.WorkingDirectory = Path.GetDirectoryName(targetFile);
                shortcut.WindowStyle = 1; // Normale vensterstijl
                shortcut.IconLocation = targetFile; // Gebruik het icoon van WoW.exe
                shortcut.Save();

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Shortcut created successfully on the desktop.\n");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed to create shortcut: {ex.Message}\n");
            }
        }
    }
}
