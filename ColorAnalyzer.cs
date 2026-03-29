using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace FileCollector
{
    public class ColorAnalyzer
    {
        public static void AnalyzeLogo()
        {
            try
            {
                const string resourceName = "FileCollector.Resources.logo.png";
                using var stream = typeof(ColorAnalyzer).Assembly.GetManifestResourceStream(resourceName);
                if (stream == null)
                {
                    Console.WriteLine($"Logo resource not found: {resourceName}");
                    return;
                }

                using (var bitmap = new Bitmap(stream))
                {
                    var colorFrequency = new Dictionary<Color, int>();
                    
                    // Анализируем все пиксели
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        for (int x = 0; x < bitmap.Width; x++)
                        {
                            Color pixel = bitmap.GetPixel(x, y);
                            
                            // Пропускаем прозрачные пиксели
                            if (pixel.A < 200) continue;
                            
                            // Округляем цвет для группировки похожих оттенков
                            Color roundedColor = Color.FromArgb(
                                (pixel.R / 10) * 10,
                                (pixel.G / 10) * 10,
                                (pixel.B / 10) * 10
                            );
                            
                            if (colorFrequency.ContainsKey(roundedColor))
                                colorFrequency[roundedColor]++;
                            else
                                colorFrequency[roundedColor] = 1;
                        }
                    }
                    
                    // Сортируем по частоте
                    var topColors = colorFrequency
                        .OrderByDescending(x => x.Value)
                        .Take(5)
                        .ToList();
                    
                    Console.WriteLine("=== Top Colors in Logo ===");
                    foreach (var entry in topColors)
                    {
                        var color = entry.Key;
                        Console.WriteLine($"Color: RGB({color.R}, {color.G}, {color.B}) | Hex: #{color.R:X2}{color.G:X2}{color.B:X2} | Count: {entry.Value}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing logo: {ex.Message}");
            }
        }
    }
}
