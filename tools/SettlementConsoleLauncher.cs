using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class SettlementConsoleLauncher
{
    [STAThread]
    private static int Main()
    {
        string appDir = AppDomain.CurrentDomain.BaseDirectory;
        string scriptPath = Path.Combine(appDir, "run-settlement-console.ps1");

        if (!File.Exists(scriptPath))
        {
            MessageBox.Show(
                "Cannot find run-settlement-console.ps1 next to this EXE.",
                "Settlement Console",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }

        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "powershell.exe";
            startInfo.Arguments = "-NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " + Quote(scriptPath);
            startInfo.WorkingDirectory = appDir;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.EnvironmentVariables["SETTLEMENT_CONSOLE_LAUNCHED_BY_EXE"] = "1";

            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();
                if (process.ExitCode != 0)
                {
                    MessageBox.Show(
                        "Settlement console failed to start. Please check startup-error-log.txt.",
                        "Settlement Console",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return process.ExitCode;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                ex.Message,
                "Settlement Console",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
