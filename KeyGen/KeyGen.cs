using System;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// Developer key generator for Excel Utilities Suite.
/// Usage: KeyGen.exe [MachineId]
/// MachineId is shown in Suite tab -> Activate License dialog on the user's PC.
/// </summary>
class KeyGen
{
    // Must match RealLicenseService._salt exactly
    private static readonly byte[] Salt =
        Encoding.UTF8.GetBytes("EUS-2024-k9mP#xQ7vR2nZ");

    static void Main(string[] args)
    {
        Console.WriteLine("=== Excel Utilities Suite — Key Generator ===");
        Console.WriteLine();

        string machineId;
        if (args.Length > 0)
        {
            machineId = args[0].Trim().ToUpperInvariant();
            Console.WriteLine("Machine ID: " + machineId);
        }
        else
        {
            Console.Write("Enter Machine ID (from Activate License dialog): ");
            machineId = (Console.ReadLine() ?? "").Trim().ToUpperInvariant();
        }

        if (machineId.Length == 0)
        {
            Console.WriteLine("ERROR: Machine ID cannot be empty.");
            Pause(); return;
        }

        string key = GenerateKey(machineId);

        Console.WriteLine();
        Console.WriteLine("Product key for this machine:");
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("  " + key);
        Console.ResetColor();
        Console.WriteLine();
        Console.WriteLine("User enters this in: Suite tab -> Activate License -> Product key");
        Console.WriteLine();
        Console.Write("Generate another key for the same machine? (y/n): ");
        if ((Console.ReadLine() ?? "").Trim().ToLower() == "y")
            Main(new[] { machineId });
        else
            Pause();
    }

    static string GenerateKey(string machineId)
    {
        byte[] rand = new byte[6];
        using (var rng = new RNGCryptoServiceProvider())
            rng.GetBytes(rand);

        string body = BitConverter.ToString(rand).Replace("-", "");
        string msg  = (machineId + body).ToUpperInvariant();

        using (var hmac = new HMACSHA256(Salt))
        {
            byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(msg));
            string tag  = BitConverter.ToString(hash, 0, 4).Replace("-", "");
            string full = body + tag;
            return full.Substring(0,5)+"-"+full.Substring(5,5)+"-"+
                   full.Substring(10,5)+"-"+full.Substring(15,5);
        }
    }

    static void Pause()
    {
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
