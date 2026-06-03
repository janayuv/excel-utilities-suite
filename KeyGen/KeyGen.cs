using System;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// Developer key generator for Excel Utilities Suite.
/// Usage: KeyGen.exe [MachineId] --salt YOUR_HMAC_SECRET
/// The salt must match LicenseSalt.Value in the add-in.
/// MachineId is shown in Suite tab -> Activate License dialog.
/// </summary>
class KeyGen
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== Excel Utilities Suite — Key Generator ===");
        Console.WriteLine();

        // Parse --salt argument
        string salt = null;
        string machineId = null;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--salt" && i + 1 < args.Length)
                salt = args[++i];
            else if (machineId == null)
                machineId = args[i].Trim().ToUpperInvariant();
        }

        if (string.IsNullOrEmpty(salt))
        {
            Console.Write("Enter HMAC salt (from LicenseSalt.cs): ");
            salt = Console.ReadLine() ?? "";
        }
        if (string.IsNullOrEmpty(salt))
        {
            Console.WriteLine("ERROR: Salt cannot be empty.");
            Pause(); return;
        }

        if (string.IsNullOrEmpty(machineId))
        {
            Console.Write("Enter Machine ID (from Activate License dialog): ");
            machineId = (Console.ReadLine() ?? "").Trim().ToUpperInvariant();
        }
        if (machineId.Length == 0)
        {
            Console.WriteLine("ERROR: Machine ID cannot be empty.");
            Pause(); return;
        }

        string key = GenerateKey(machineId, salt);

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
        {
            // Pass salt via args on recursive call
            Main(new[] { machineId, "--salt", salt });
        }
        else
            Pause();
    }

    static string GenerateKey(string machineId, string saltStr)
    {
        byte[] saltBytes = System.Text.Encoding.UTF8.GetBytes(saltStr);
        byte[] rand = new byte[6];
        using (var rng = new RNGCryptoServiceProvider())
            rng.GetBytes(rand);

        string body = BitConverter.ToString(rand).Replace("-", "");
        string msg  = (machineId + body).ToUpperInvariant();

        using (var hmac = new HMACSHA256(saltBytes))
        {
            byte[] hash = hmac.ComputeHash(System.Text.Encoding.UTF8.GetBytes(msg));
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
