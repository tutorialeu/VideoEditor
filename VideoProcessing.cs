using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using NReco.VideoConverter;
using MediaToolkit;
using MediaToolkit.Options;
using System;
using MediaToolkit.Model;
using System.Drawing;
using NAudio.Wave;
using NAudio.WaveFormRenderer;
using WMPLib;
using System.Linq;
using System.Diagnostics;

namespace VideoEditor
{
    public class VideoProcessing
    {

        public static string FileNameWithoutExtension(string fileName)
        {
            return Path.GetFileNameWithoutExtension(fileName);
        }

        public static void ConvertMp4ToMp3(string inputFile, string outputFile)
        {
            if (File.Exists(outputFile))
            {
                MessageBox.Show(outputFile + " already exists!");
                return;
            }
            var ConvertVideo = new FFMpegConverter();
            var dirPath = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }
            if (!File.Exists(outputFile))
            {
                ConvertVideo.ConvertMedia(inputFile, outputFile, "mp3");
            }
            else
            {
                MessageBox.Show("File: " + outputFile + "already exist!");
            }
        }
        public static void ConvertFileToWAV(string inputFile, int duration, string outputFile)
        {
            if (!File.Exists(inputFile))
            {
                return;
            }
            if (File.Exists(outputFile))
            {
                var player = new WindowsMediaPlayer();
                var clip = player.newMedia(outputFile).duration;
                if (duration != clip)
                {
                    File.Delete(outputFile);
                }
                else
                {
                    return;
                }
            }
            var dirPath = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }
            if (!File.Exists(outputFile))
            {
                try
                {
                    var ConvertVideo = new FFMpegConverter();
                    ConvertVideo.ConvertMedia(inputFile, outputFile, "wav");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Audio extraction failed: " + ex.Message);
                }
            }
        }
        public static Image GetVideoTumbnail(string inputFile, float frame, string outputPath)
        {
            if (!File.Exists(inputFile))
            {
                return null;
            }
            try
            {
                string fileNameNoExt = Path.GetFileNameWithoutExtension(inputFile);
                // Sanitize the filename for use in output path (remove special chars)
                string safeFileName = string.Join("_", fileNameNoExt.Split(Path.GetInvalidFileNameChars()));
                string outputFile = Path.Combine(outputPath, safeFileName + "_" + frame + ".jpeg");
                
                if (File.Exists(outputFile))
                {
                    return Image.FromFile(outputFile);
                }
                
                var ConvertVideo = new FFMpegConverter();
                ConvertVideo.GetVideoThumbnail(inputFile, outputFile, frame);
                
                if (File.Exists(outputFile))
                {
                    return Image.FromFile(outputFile);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Thumbnail generation failed at frame " + frame + ": " + ex.Message);
            }
            return null;
        }
        public static void ConcatVideo(string[] inputFile, string outputFile, Action<double> progressCallback = null)
        {
            if (File.Exists(outputFile))
            {
                MessageBox.Show(outputFile + " already exists!");
                return;
            }
            var ConvertVideo = new FFMpegConverter();
            if (progressCallback != null)
            {
                ConvertVideo.ConvertProgress += (sender, e) =>
                {
                    double progress = (double)e.Processed.TotalSeconds / e.TotalDuration.TotalSeconds;
                    progressCallback(progress);
                };
            }
            ConcatSettings cs = new ConcatSettings();

            ConvertVideo.ConcatMedia(inputFile, outputFile, "mp4", cs);
        }

        public static void CutAndConcat(List<Item> videoItems, string outputFile, Action<double> progressCallback = null)
        {
            if (File.Exists(outputFile))
            {
                MessageBox.Show(outputFile + " already exists!");
                return;
            }

            var tempFiles = new List<string>();
            var converter = new FFMpegConverter();
            string tempDir = Path.Combine(Path.GetTempPath(), "VideoEditorTemp");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);

            try
            {
                int totalSteps = videoItems.Count + 1; // Cuts + 1 Merge
                int currentStep = 0;

                // 1. Cut segments
                foreach (var item in videoItems)
                {
                    string tempFile = Path.Combine(tempDir, Guid.NewGuid().ToString() + ".mp4");
                    tempFiles.Add(tempFile);

                    // We use ConvertMedia with specific settings to cut
                    // startVideo is the start offset in the SOURCE file
                    // duration is the length on the timeline
                    converter.ConvertMedia(item.path, null, tempFile, "mp4", new ConvertSettings()
                    {
                        Seek = (float)item.startVideo,
                        MaxDuration = (float)item.duration
                    });

                    currentStep++;
                    progressCallback?.Invoke((double)currentStep / totalSteps);
                }

                // 2. Concat segments
                ConcatSettings cs = new ConcatSettings();
                converter.ConvertProgress += (sender, e) =>
                {
                    double mergeProgress = (double)e.Processed.TotalSeconds / e.TotalDuration.TotalSeconds;
                    // For the merge phase, we map progress from [currentStep/totalSteps] to 1.0
                    double overallProgress = ((double)currentStep / totalSteps) + (mergeProgress * (1.0 / totalSteps));
                    progressCallback?.Invoke(Math.Min(0.99, overallProgress));
                };

                converter.ConcatMedia(tempFiles.ToArray(), outputFile, "mp4", cs);
                progressCallback?.Invoke(1.0);
            }
            finally
            {
                // Cleanup
                foreach (var file in tempFiles)
                {
                    try { if (File.Exists(file)) File.Delete(file); } catch { }
                }
            }
        }
        public static void CutVideo(string input, string output, TimeSpan start, TimeSpan duration)
        {
            var inputFile = new MediaFile { Filename = @input };
            var outputFile = new MediaFile { Filename = @output };

            using (var engine = new Engine())
            {
                var options = new ConversionOptions();
                options.CutMedia(start, duration);
                engine.Convert(inputFile, outputFile, options);
            }
        }

        public static Image CreateWaveImage(string path, int duration)
        {
            // Scale waveform width with duration for better resolution
            int waveWidth = Math.Max(640, duration * 10); // ~10 pixels per second
            if (waveWidth > 8000) waveWidth = 8000; // Cap to avoid memory issues
            int waveHeight = 70; // Fill most of the track height (80px)

            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return new Bitmap(waveWidth, waveHeight);
            }

            try
            {
                var myRendererSettings = new StandardWaveFormRendererSettings();
                myRendererSettings.Width = waveWidth;
                myRendererSettings.TopHeight = waveHeight / 2;
                myRendererSettings.BottomHeight = waveHeight / 2;
                myRendererSettings.BackgroundColor = Color.Transparent;

                var renderer = new WaveFormRenderer();
                using (var reader = new AudioFileReader(path))
                {
                    // Use a wrapper to ensure block-aligned reads
                    using (var alignedReader = new BlockAlignedStream(reader))
                    {
                        var image = renderer.Render(alignedReader, myRendererSettings);
                        return image;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Waveform rendering failed: " + ex.Message);
                
                Bitmap bmp = new Bitmap(waveWidth, waveHeight);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.FromArgb(50, Color.Gray));
                    g.DrawString("Audio Waveform unavailable", SystemFonts.DefaultFont, Brushes.White, 10, 10);
                }
                return bmp;
            }
        }

        /// <summary>
        /// A wrapper for WaveStream that ensures all reads are block-aligned.
        /// Useful for renderers that might request unaligned byte counts.
        /// </summary>
        private class BlockAlignedStream : WaveStream
        {
            private readonly WaveStream sourceStream;
            public BlockAlignedStream(WaveStream sourceStream) { this.sourceStream = sourceStream; }
            public override WaveFormat WaveFormat => sourceStream.WaveFormat;
            public override long Length => sourceStream.Length;
            public override long Position
            {
                get => sourceStream.Position;
                set => sourceStream.Position = (value / sourceStream.WaveFormat.BlockAlign) * sourceStream.WaveFormat.BlockAlign;
            }
            public override int Read(byte[] buffer, int offset, int count)
            {
                int alignedCount = (count / sourceStream.WaveFormat.BlockAlign) * sourceStream.WaveFormat.BlockAlign;
                if (alignedCount == 0) return 0;
                return sourceStream.Read(buffer, offset, alignedCount);
            }
            protected override void Dispose(bool disposing)
            {
                if (disposing) sourceStream.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
