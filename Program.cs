using System;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Speech.AudioFormat;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAudio.Wave;
using NAudio.Lame;

// WinRT / OneCore - use full namespace to avoid ambiguity
using WinRTSpeechSynthesizer = Windows.Media.SpeechSynthesis.SpeechSynthesizer;
using Windows.Media.SpeechSynthesis;
using Windows.Storage.Streams;

namespace TTSGui
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        enum EngineMode { Auto, SAPI, WinRT }

        private TextBox txtInput;
        private ComboBox cboEngine;
        private ComboBox cboLang;
        private ComboBox cboVoice;
        private ComboBox cboPitch;
        private TrackBar trkRate;
        private Label lblPitch, lblRate, lblStatus;
        private Label lblTitle, lblEngine, lblLang, lblVoice;
        private Panel pnlHeader, pnlControls, pnlActions;
        private Button btnRefreshVoices, btnPreview, btnStop, btnSaveMp3;

        private System.Speech.Synthesis.SpeechSynthesizer sapiSynth = new System.Speech.Synthesis.SpeechSynthesizer();
        private WaveOutEvent? waveOut;
        private AudioFileReader? currentReader;

        // Modern Color Scheme - Light Theme
        private static readonly System.Drawing.Color PrimaryColor = System.Drawing.Color.FromArgb(79, 70, 229); // Indigo
        private static readonly System.Drawing.Color SecondaryColor = System.Drawing.Color.FromArgb(99, 102, 241); // Lighter Indigo
        private static readonly System.Drawing.Color AccentColor = System.Drawing.Color.FromArgb(139, 92, 246); // Purple
        private static readonly System.Drawing.Color BackgroundColor = System.Drawing.Color.FromArgb(249, 250, 251); // Light Gray
        private static readonly System.Drawing.Color SurfaceColor = System.Drawing.Color.White;
        private static readonly System.Drawing.Color TextColor = System.Drawing.Color.FromArgb(31, 41, 55); // Dark Gray
        private static readonly System.Drawing.Color TextSecondaryColor = System.Drawing.Color.FromArgb(107, 114, 128); // Medium Gray
        private static readonly System.Drawing.Color BorderColor = System.Drawing.Color.FromArgb(229, 231, 235); // Light Border
        private static readonly System.Drawing.Color SuccessColor = System.Drawing.Color.FromArgb(16, 185, 129); // Green
        private static readonly System.Drawing.Color DangerColor = System.Drawing.Color.FromArgb(239, 68, 68); // Red

        public MainForm()
        {
            Text = "🎙️ Text to Speech - Thai & English";
            Width = 1000;
            Height = 750;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = BackgroundColor;
            Font = new System.Drawing.Font("Segoe UI", 9.5f);

            // HEADER PANEL
            pnlHeader = new Panel
            {
                Left = 0, Top = 0, Width = 1000, Height = 80,
                BackColor = PrimaryColor
            };

            lblTitle = new Label
            {
                Text = "🎙️ Text to Speech",
                Left = 25, Top = 15, Width = 400, Height = 35,
                Font = new System.Drawing.Font("Segoe UI", 18f, System.Drawing.FontStyle.Bold),
                ForeColor = System.Drawing.Color.White,
                BackColor = System.Drawing.Color.Transparent
            };

            var lblSubtitle = new Label
            {
                Text = "Convert text to natural speech in Thai & English",
                Left = 25, Top = 50, Width = 500, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f),
                ForeColor = System.Drawing.Color.FromArgb(224, 231, 255),
                BackColor = System.Drawing.Color.Transparent
            };

            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblSubtitle });

            // INPUT AREA
            var lblInputTitle = new Label
            {
                Text = "📝 Input Text",
                Left = 30, Top = 100, Width = 200, Height = 25,
                Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold),
                ForeColor = TextColor
            };

            txtInput = new TextBox
            {
                Multiline = true, ScrollBars = ScrollBars.Vertical,
                Left = 30, Top = 130, Width = 920, Height = 280,
                Font = new System.Drawing.Font("Segoe UI", 11f),
                BackColor = SurfaceColor,
                ForeColor = TextColor,
                BorderStyle = BorderStyle.FixedSingle
            };

            // CONTROLS PANEL
            pnlControls = new Panel
            {
                Left = 20, Top = 430, Width = 940, Height = 180,
                BackColor = SurfaceColor,
                BorderStyle = BorderStyle.FixedSingle
            };

            int x = 20; int y = 20; int gap = 15;

            // Engine Selection
            lblEngine = new Label
            {
                Text = "⚙️ Engine",
                Left = x, Top = y, Width = 170, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = TextSecondaryColor
            };

            cboEngine = new ComboBox
            {
                Left = x, Top = y + 22, Width = 170, Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                BackColor = SurfaceColor,
                ForeColor = TextColor
            };
            cboEngine.Items.AddRange(new object[] { "Auto", "SAPI (System.Speech)", "WinRT (OneCore)" });
            cboEngine.SelectedIndex = 0;
            x += 170 + gap;

            // Language Selection
            lblLang = new Label
            {
                Text = "🌐 Language",
                Left = x, Top = y, Width = 160, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = TextSecondaryColor
            };

            cboLang = new ComboBox
            {
                Left = x, Top = y + 22, Width = 160, Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                BackColor = SurfaceColor,
                ForeColor = TextColor
            };
            cboLang.Items.AddRange(new object[] { "Auto", "Thai (th-TH)", "English (en-US)" });
            cboLang.SelectedIndex = 0;
            x += 160 + gap;

            // Voice Selection
            lblVoice = new Label
            {
                Text = "🎤 Voice",
                Left = x, Top = y, Width = 360, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = TextSecondaryColor
            };

            cboVoice = new ComboBox
            {
                Left = x, Top = y + 22, Width = 360, Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                BackColor = SurfaceColor,
                ForeColor = TextColor
            };
            x += 360 + gap;

            btnRefreshVoices = new Button
            {
                Left = x, Top = y + 21, Width = 110, Height = 30,
                Text = "↻ Refresh",
                Font = new System.Drawing.Font("Segoe UI", 9f),
                BackColor = SecondaryColor,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnRefreshVoices.FlatAppearance.BorderSize = 0;
            x = 20; y += 70;

            // Pitch Control
            lblPitch = new Label
            {
                Text = "🎵 Pitch",
                Left = x, Top = y, Width = 180, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = TextSecondaryColor
            };

            cboPitch = new ComboBox
            {
                Left = x, Top = y + 22, Width = 180, Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 9.5f),
                BackColor = SurfaceColor,
                ForeColor = TextColor
            };
            cboPitch.Items.AddRange(new object[] { 
                "Extra Low (-40%)",
                "Very Low (-30%)", 
                "Low (-20%)", 
                "Slightly Low (-10%)",
                "Normal (0%)", 
                "Slightly High (+10%)",
                "High (+20%)", 
                "Very High (+30%)",
                "Extra High (+40%)"
            });
            cboPitch.SelectedIndex = 4; // Normal
            x += 180 + gap;

            // Speed Control
            var lblSpeed = new Label
            {
                Text = "⚡ Speed",
                Left = x, Top = y, Width = 250, Height = 20,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold),
                ForeColor = TextSecondaryColor
            };

            trkRate = new TrackBar
            {
                Left = x, Top = y + 15, Width = 250, Height = 45,
                Minimum = -10, Maximum = 10, TickFrequency = 2, Value = 0,
                BackColor = SurfaceColor
            };

            lblRate = new Label
            {
                Left = x + 260, Top = y + 24, Width = 100, Height = 25,
                Text = "0",
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            pnlControls.Controls.AddRange(new Control[] {
                lblEngine, cboEngine, lblLang, cboLang, lblVoice, cboVoice, btnRefreshVoices,
                lblPitch, cboPitch, lblSpeed, trkRate, lblRate
            });

            // ACTIONS PANEL
            pnlActions = new Panel
            {
                Left = 20, Top = 625, Width = 940, Height = 70,
                BackColor = System.Drawing.Color.Transparent
            };

            x = 20; y = 10;

            btnPreview = new Button
            {
                Left = x, Top = y, Width = 200, Height = 50,
                Text = "▶  Preview",
                Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold),
                BackColor = SuccessColor,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnPreview.FlatAppearance.BorderSize = 0;
            x += 220;

            btnStop = new Button
            {
                Left = x, Top = y, Width = 150, Height = 50,
                Text = "⏹  Stop",
                Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold),
                BackColor = DangerColor,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnStop.FlatAppearance.BorderSize = 0;
            x += 170;

            btnSaveMp3 = new Button
            {
                Left = x, Top = y, Width = 220, Height = 50,
                Text = "💾  Save as MP3",
                Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold),
                BackColor = AccentColor,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSaveMp3.FlatAppearance.BorderSize = 0;

            pnlActions.Controls.AddRange(new Control[] { btnPreview, btnStop, btnSaveMp3 });

            // STATUS BAR
            lblStatus = new Label
            {
                Left = 30, Top = 705, Width = 920, Height = 25,
                Text = "✓ Ready to convert text to speech",
                Font = new System.Drawing.Font("Segoe UI", 9f),
                ForeColor = TextSecondaryColor,
                AutoEllipsis = true
            };

            Controls.AddRange(new Control[] {
                pnlHeader, lblInputTitle, txtInput, pnlControls, pnlActions, lblStatus
            });

            // Add hover effects
            AddHoverEffect(btnRefreshVoices, SecondaryColor);
            AddHoverEffect(btnPreview, SuccessColor);
            AddHoverEffect(btnStop, DangerColor);
            AddHoverEffect(btnSaveMp3, AccentColor);

            // events
            trkRate.ValueChanged += (s, e) => lblRate.Text = $"Speed: {trkRate.Value}";
            btnRefreshVoices.Click += (s, e) => PopulateVoices();
            cboEngine.SelectedIndexChanged += (s, e) => PopulateVoices();
            cboLang.SelectedIndexChanged += (s, e) => PopulateVoices();
            btnPreview.Click += async (s, e) => await PreviewAsync();
            btnStop.Click += (s, e) => StopPlayback();
            btnSaveMp3.Click += async (s, e) => await SaveMp3Async();

            PopulateVoices();
        }

        // ---------- UI Utilities ----------
        private void AddHoverEffect(Button btn, System.Drawing.Color baseColor)
        {
            var hoverColor = System.Drawing.Color.FromArgb(
                Math.Min(baseColor.R + 20, 255),
                Math.Min(baseColor.G + 20, 255),
                Math.Min(baseColor.B + 20, 255)
            );

            btn.MouseEnter += (s, e) => btn.BackColor = hoverColor;
            btn.MouseLeave += (s, e) => btn.BackColor = baseColor;
        }

        // ---------- Utility ----------
        private static bool ContainsThai(string text) => Regex.IsMatch(text, "[\u0E00-\u0E7F]");

        private EngineMode SelectedEngine =>
            cboEngine.SelectedIndex switch { 1 => EngineMode.SAPI, 2 => EngineMode.WinRT, _ => EngineMode.Auto };

        private string SelectedLang(string text)
        {
            if (cboLang.SelectedIndex == 1) return "th-TH";
            if (cboLang.SelectedIndex == 2) return "en-US";
            return ContainsThai(text) ? "th-TH" : "en-US";
        }

        private string BuildSsml(string text, int rateSteps, string pitchPreset, string lang)
        {
            int ratePercent = rateSteps * 10; // -100..100
            string rateStr = (ratePercent >= 0 ? "+" : "") + ratePercent + "%";
            string pitchStr = ParsePitchValue(pitchPreset);

            return $@"<speak version='1.0' xml:lang='{lang}'>
                        <prosody rate='{rateStr}' pitch='{pitchStr}'>
                          {System.Security.SecurityElement.Escape(text)}
                        </prosody>
                      </speak>";
        }

        private string ParsePitchValue(string pitchPreset)
        {
            // Extract percentage from preset string like "Low (-20%)" or "Normal (0%)"
            if (pitchPreset.Contains("(-40%)")) return "-40%";
            if (pitchPreset.Contains("(-30%)")) return "-30%";
            if (pitchPreset.Contains("(-20%)")) return "-20%";
            if (pitchPreset.Contains("(-10%)")) return "-10%";
            if (pitchPreset.Contains("(0%)")) return "0%";
            if (pitchPreset.Contains("(+10%)")) return "+10%";
            if (pitchPreset.Contains("(+20%)")) return "+20%";
            if (pitchPreset.Contains("(+30%)")) return "+30%";
            if (pitchPreset.Contains("(+40%)")) return "+40%";
            
            // Fallback for old format
            if (pitchPreset.Contains("Low")) return "-20%";
            if (pitchPreset.Contains("High")) return "+20%";
            return "0%";
        }

        // ---------- Voice listing ----------
        private void PopulateVoices()
        {
            cboVoice.Items.Clear();

            var lang = SelectedLang(txtInput.Text);
            if (SelectedEngine == EngineMode.SAPI || SelectedEngine == EngineMode.Auto)
            {
                var voices = sapiSynth.GetInstalledVoices()
                                      .Where(v => v.Enabled && v.VoiceInfo?.Culture?.Name == lang)
                                      .Select(v => $"SAPI|{v.VoiceInfo.Name}|{v.VoiceInfo.Culture}")
                                      .OrderBy(s => s)
                                      .ToList();
                foreach (var v in voices) cboVoice.Items.Add(v);
            }

            if (SelectedEngine == EngineMode.WinRT || SelectedEngine == EngineMode.Auto)
            {
                var winrtVoices = WinRTSpeechSynthesizer.AllVoices
                                    .Where(v => v.Language == lang)
                                    .Select(v => $"WINRT|{v.DisplayName}|{v.Language}")
                                    .OrderBy(s => s)
                                    .ToList();
                foreach (var v in winrtVoices) cboVoice.Items.Add(v);
            }

            if (cboVoice.Items.Count == 0)
                cboVoice.Items.Add("⚠ No voices for selected options");

            cboVoice.SelectedIndex = 0;
            lblStatus.Text = $"✓ Found {cboVoice.Items.Count} voice(s) for {lang}";
        }

        private bool HasSapiVoice(string culture) =>
            sapiSynth.GetInstalledVoices().Any(v => v.Enabled && v.VoiceInfo.Culture.Name.Equals(culture, StringComparison.OrdinalIgnoreCase));

        private bool HasWinRtVoice(string culture) =>
            WinRTSpeechSynthesizer.AllVoices.Any(v => v.Language.Equals(culture, StringComparison.OrdinalIgnoreCase));

        // ---------- Preview / Playback ----------
        private void StopPlayback()
        {
            try
            {
                waveOut?.Stop();
                waveOut?.Dispose(); waveOut = null;
                currentReader?.Dispose(); currentReader = null;
                lblStatus.Text = "⏹ Playback stopped";
            }
            catch { }
        }

        private async Task PreviewAsync()
        {
            StopPlayback();

            var text = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(text)) { MessageBox.Show("กรุณาพิมพ์ข้อความ", "Info"); return; }

            var lang = SelectedLang(text);
            var engine = ResolveEngineFor(lang);

            lblStatus.Text = $"▶ Previewing with {engine} ({lang})...";

            try
            {
                string tempWav = await SynthesizeToWavAsync(text, lang, engine,
                    trkRate.Value, cboPitch.SelectedItem?.ToString() ?? "Medium", fromPreview: true);

                currentReader = new AudioFileReader(tempWav);
                waveOut = new WaveOutEvent();
                waveOut.Init(currentReader);
                waveOut.Play();

                waveOut.PlaybackStopped += (s, e) =>
                {
                    try { if (File.Exists(tempWav)) File.Delete(tempWav); } catch { }
                    currentReader?.Dispose(); currentReader = null;
                    waveOut?.Dispose(); waveOut = null;
                };
            }
            catch (Exception ex)
            {
                lblStatus.Text = "❌ Preview failed";
                MessageBox.Show(ex.Message, "Preview Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---------- Save MP3 ----------
        private async Task SaveMp3Async()
        {
            var text = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(text)) { MessageBox.Show("กรุณาพิมพ์ข้อความ", "Info"); return; }

            using var sfd = new SaveFileDialog { Filter = "MP3 Audio (*.mp3)|*.mp3", FileName = "tts.mp3" };
            if (sfd.ShowDialog() != DialogResult.OK) return;

            var lang = SelectedLang(text);
            var engine = ResolveEngineFor(lang);

            lblStatus.Text = $"💾 Converting to MP3 with {engine}...";

            string wavPath = "";
            try
            {
                wavPath = await SynthesizeToWavAsync(text, lang, engine,
                    trkRate.Value, cboPitch.SelectedItem?.ToString() ?? "Medium", fromPreview: false);

                using var reader = new AudioFileReader(wavPath);
                using var mp3 = new LameMP3FileWriter(sfd.FileName, reader.WaveFormat, LAMEPreset.VBR_90);
                reader.CopyTo(mp3);

                lblStatus.Text = $"✅ Saved successfully: {System.IO.Path.GetFileName(sfd.FileName)}";
                MessageBox.Show("บันทึก MP3 สำเร็จ! 🎉", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "❌ Failed to save MP3";
                MessageBox.Show(ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try { if (!string.IsNullOrEmpty(wavPath) && File.Exists(wavPath)) File.Delete(wavPath); } catch { }
            }
        }

        private EngineMode ResolveEngineFor(string lang)
        {
            var requested = SelectedEngine;
            if (requested == EngineMode.SAPI) return EngineMode.SAPI;
            if (requested == EngineMode.WinRT) return EngineMode.WinRT;

            // Auto: SAPI → WinRT
            if (HasSapiVoice(lang)) return EngineMode.SAPI;
            if (HasWinRtVoice(lang)) return EngineMode.WinRT;
            // ถ้าไม่มีทั้งคู่: ให้ลอง SAPI (จะ throw ชัดเจน)
            return EngineMode.SAPI;
        }

        private async Task<string> SynthesizeToWavAsync(string text, string lang, EngineMode engine, int rateSteps, string pitchPreset, bool fromPreview)
        {
            return engine switch
            {
                EngineMode.SAPI => SynthesizeWithSapiToWav(text, lang, rateSteps, pitchPreset),
                EngineMode.WinRT => await SynthesizeWithWinRtToWavAsync(text, lang, rateSteps, pitchPreset),
                _ => throw new InvalidOperationException("Unknown engine")
            };
        }

        // ---------- SAPI ----------
        private string SynthesizeWithSapiToWav(string text, string lang, int rateSteps, string pitchPreset)
        {
            var voiceName = ParseSelectedVoice(cboVoice.SelectedItem?.ToString(), pickEngine: EngineMode.SAPI);
            if (!string.IsNullOrEmpty(voiceName))
            {
                try { sapiSynth.SelectVoice(voiceName); } catch { /* ถ้าเลือกไม่ได้จะใช้ default */ }
            }
            else
            {
                // เลือกเสียง SAPI ตามภาษาอัตโนมัติ
                var v = sapiSynth.GetInstalledVoices()
                                 .FirstOrDefault(vv => vv.Enabled && vv.VoiceInfo.Culture.Name == lang);
                if (v != null) sapiSynth.SelectVoice(v.VoiceInfo.Name);
            }

            string ssml = BuildSsml(text, rateSteps, pitchPreset, lang);
            string wavPath = Path.Combine(Path.GetTempPath(), $"tts_sapi_{Guid.NewGuid()}.wav");

            // Mono 44.1k
            sapiSynth.SetOutputToWaveFile(wavPath, new SpeechAudioFormatInfo(44100, AudioBitsPerSample.Sixteen, AudioChannel.Mono));
            sapiSynth.Speak(new Prompt(ssml, SynthesisTextFormat.Ssml));
            sapiSynth.SetOutputToDefaultAudioDevice();

            return wavPath;
        }

        // ---------- WinRT ----------
        private async Task<string> SynthesizeWithWinRtToWavAsync(string text, string lang, int rateSteps, string pitchPreset)
        {
            var synth = new WinRTSpeechSynthesizer();

            var voiceIdOrName = ParseSelectedVoice(cboVoice.SelectedItem?.ToString(), pickEngine: EngineMode.WinRT);
            VoiceInformation? chosen = null;
            if (!string.IsNullOrEmpty(voiceIdOrName))
            {
                // หาแบบ DisplayName ก่อน
                chosen = WinRTSpeechSynthesizer.AllVoices.FirstOrDefault(v => v.DisplayName.Equals(voiceIdOrName, StringComparison.OrdinalIgnoreCase));
                // ไม่เจอ ลองหาแบบ Id
                if (chosen == null) chosen = WinRTSpeechSynthesizer.AllVoices.FirstOrDefault(v => v.Id == voiceIdOrName);
            }
            if (chosen == null)
                chosen = WinRTSpeechSynthesizer.AllVoices.FirstOrDefault(v => v.Language == lang);
            if (chosen != null) synth.Voice = chosen;

            int ratePercent = rateSteps * 10;
            string rateStr = (ratePercent >= 0 ? "+" : "") + ratePercent + "%";
            string pitchStr = ParsePitchValue(pitchPreset);
            string ssml = $@"<speak version='1.0' xml:lang='{lang}'>
                                <prosody rate='{rateStr}' pitch='{pitchStr}'>
                                  {System.Security.SecurityElement.Escape(text)}
                                </prosody>
                             </speak>";

            var stream = await synth.SynthesizeSsmlToStreamAsync(ssml);
            string wavPath = Path.Combine(Path.GetTempPath(), $"tts_winrt_{Guid.NewGuid()}.wav");
            using (var file = File.Create(wavPath))
            {
                stream.AsStreamForRead().CopyTo(file);
            }
            return wavPath;
        }

        // parse "SAPI|<name>|<culture>" or "WINRT|<display>|<lang>"
        private static string? ParseSelectedVoice(string? itemText, EngineMode pickEngine)
        {
            if (string.IsNullOrWhiteSpace(itemText)) return null;
            var parts = itemText.Split('|');
            if (parts.Length < 3) return null;
            if (pickEngine == EngineMode.SAPI && parts[0] == "SAPI") return parts[1];
            if (pickEngine == EngineMode.WinRT && parts[0] == "WINRT") return parts[1];
            return null;
        }
    }
}
