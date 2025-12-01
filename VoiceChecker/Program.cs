using System;
using System.Linq;
using System.Speech.Synthesis;
using Windows.Media.SpeechSynthesis;

class CheckVoices
{
    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        
        Console.WriteLine("=== SAPI Voices (System.Speech) ===");
        var sapiSynth = new System.Speech.Synthesis.SpeechSynthesizer();
        var sapiVoices = sapiSynth.GetInstalledVoices();
        foreach (var voice in sapiVoices)
        {
            if (voice.Enabled)
            {
                var info = voice.VoiceInfo;
                Console.WriteLine($"- {info.Name} ({info.Culture.Name}) - {info.Gender}");
            }
        }
        
        Console.WriteLine("\n=== WinRT Voices (Windows.Media.SpeechSynthesis) ===");
        var winrtVoices = Windows.Media.SpeechSynthesis.SpeechSynthesizer.AllVoices;
        foreach (var voice in winrtVoices)
        {
            Console.WriteLine($"- {voice.DisplayName} ({voice.Language}) - {voice.Gender}");
        }
        
        Console.WriteLine("\n=== Thai Voices Available? ===");
        var hasSapiThai = sapiVoices.Any(v => v.Enabled && v.VoiceInfo.Culture.Name.StartsWith("th"));
        var hasWinRtThai = winrtVoices.Any(v => v.Language.StartsWith("th"));
        
        Console.WriteLine($"SAPI Thai: {(hasSapiThai ? "✓ YES" : "✗ NO")}");
        Console.WriteLine($"WinRT Thai: {(hasWinRtThai ? "✓ YES" : "✗ NO")}");
        
        if (!hasSapiThai && !hasWinRtThai)
        {
            Console.WriteLine("\n⚠️  WARNING: No Thai voices found!");
            Console.WriteLine("To add Thai voice:");
            Console.WriteLine("1. Open Windows Settings");
            Console.WriteLine("2. Go to Time & Language > Language & Region");
            Console.WriteLine("3. Click 'Add a language' and select 'Thai'");
            Console.WriteLine("4. After installing, go to Speech settings");
            Console.WriteLine("5. Download Thai speech pack");
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}
