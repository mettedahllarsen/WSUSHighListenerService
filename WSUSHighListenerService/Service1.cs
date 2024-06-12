using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WSUSHighListenerService
{
	public partial class Service1 : ServiceBase
	{

		TcpListener listener;

		public Service1()
		{
			InitializeComponent();
		}

		protected override void OnStart(string[] args)
		{
			// Start TCP listener på en bestemt port
			int port = 5001; // Eksempelport, juster efter behov
			listener = new TcpListener(IPAddress.Any, port);
			listener.Start();

			Console.WriteLine($"Server started on port {port}.");

			// Loop for at lytte på anmodninger
			while (true)
			{
				// Accept anmodning
				TcpClient client = listener.AcceptTcpClient();
				Console.WriteLine("Anmodning modtaget.");

				// Håndter anmodningen i en separat tråd
				Task.Run(() => HandleClient(client));
			}
		}

		protected override void OnStop()
		{
			listener.Stop();
			Console.WriteLine("Server stopped.");
		}

		private void HandleClient(TcpClient client)
		{
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] buffer = new byte[1024];
				int bytesRead = stream.Read(buffer, 0, buffer.Length);
				string message = Encoding.UTF8.GetString(buffer, 0, bytesRead);

				// Identificer typen af request og udfør handling
				switch (message)
				{
					case "GET_VERSION":
						// Få Windows-versionen ved hjælp af PowerShell
						string version = RunPowerShellCommand("Get-ComputerInfo -Property 'WindowsVersion'").Trim();
						byte[] responseBytes = Encoding.UTF8.GetBytes(version);
						stream.Write(responseBytes, 0, responseBytes.Length);
						break;
					case "SET_UPDATE":
						// Implementer logik for at læse opdateringsfil og udføre opdatering

						// Enheden kigger ned i en mappe på eget drev for at finde opdateringen der er blevet pushet (.msi fil?)
						string updateFilePath = GetUpdateFilePath(); // Hent stien til opdateringsfilen
						if (!File.Exists(updateFilePath))
						{
							SendResponse(stream, "Ingen opdateringsfil fundet.");
							break;
						}

						// Script der igangsætter opdateringen 

						// Fejlhåndtering

						// * Afbrudt forbindelse eller lign = rollback --> prøv igen

						break;
					default:
						// Behandling af ukendte eller fejlbehandlinger
						Console.WriteLine($"Ukendt anmodning: {message}");
						break;
				}

				client.Close();
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Fejl ved behandling af anmodning: {ex.Message}");
			}
		}

		private static string GetUpdateFilePath()
		{
			// Her skal du definere logikken til at finde stien til opdateringsfilen.
			// Du kan f.eks. bruge en fast defineret mappe eller en konfigurationsfil.
			string updateFolder = @"C:\Updates"; // Eksempel på en fast defineret mappe
			string updateFileName = "latest.msi"; // Eksempel på filnavn

			string updateFilePath = Path.Combine(updateFolder, updateFileName);
			return updateFilePath;
		}
		private string RunPowerShellCommand(string command)
		{
			using (var process = new Process())
			{
				process.StartInfo.FileName = "powershell.exe";
				process.StartInfo.Arguments = $"-Command \"{command}\"";
				process.StartInfo.UseShellExecute = false;
				process.StartInfo.RedirectStandardOutput = true;
				process.Start();

				string output = process.StandardOutput.ReadToEnd();
				process.WaitForExit();

				return output;
			}
		}

		private static void SendResponse(NetworkStream stream, string message)
		{
			byte[] responseBytes = Encoding.UTF8.GetBytes(message);
			stream.Write(responseBytes, 0, responseBytes.Length);
		}

		private void InstallMSIUpdates()
		{
			// Get the path to the update directory
			string updateDirPath = GetUpdateFilePath();

			// Find the first.msi file in the directory
			string msiFile = Directory.GetFiles(updateDirPath, "*.msi").FirstOrDefault();

			if (!string.IsNullOrEmpty(msiFile))
			{
				try
				{
					ProcessStartInfo startInfo = new ProcessStartInfo
					{
						FileName = msiFile,
						UseShellExecute = true,
						Verb = "runas"
					};

					Process installationProcess = Process.Start(startInfo);
					installationProcess.WaitForExit();

					if (installationProcess.ExitCode != 0) // Check if installation failed
					{
						// Initiate rollback using msiexec
						ProcessStartInfo rollbackInfo = new ProcessStartInfo
						{
							FileName = "msiexec.exe",
							Arguments = $"/x {msiFile} /quiet", // Uninstall quietly
							UseShellExecute = true,
							Verb = "runas"
						};

						Process.Start(rollbackInfo);
						Console.WriteLine($"Rollback udført for opdatering: {Path.GetFileName(msiFile)}");
					}
					else
					{
						Console.WriteLine($"Installeret opdatering: {Path.GetFileName(msiFile)}");
						// File.Delete(msiFile); // Optionally delete the .msi file
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Fejl ved installation af {Path.GetFileName(msiFile)}: {ex.Message}");
					// Du kan forsøge en rollback her, hvis installationen fejlede på grund af en exception
					Process.Start(new ProcessStartInfo("msiexec.exe", $"/x {msiFile} /quiet") { UseShellExecute = true, Verb = "runas" });
				}
			}
		}
	}
}
