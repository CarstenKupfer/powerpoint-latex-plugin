using System;
using WpfMath;
using WpfMath.Rendering;

namespace PowerPointLatexPlugin
{
    internal static class LatexRenderer
    {
        public static string RenderToPng(string latex)
        {
            var parser =
                WpfMath.Parsers.WpfTeXFormulaParser.Instance;

            var formula =
                parser.Parse(latex);

            var pngBytes =
                formula.RenderToPng(
                    30.0,
                    0.0,
                    0.0,
                    "Arial"
                );

            string filePath =
                System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "PowerPointLatexPlugin_formula.png"
                );

            System.IO.File.WriteAllBytes(
                filePath,
                pngBytes
            );

            return filePath;
        }
        public static string RenderToSvg(string latex, double opacity = 1.0)
        {
            opacity = Math.Max(0.0, Math.Min(1.0, opacity));

            string opacityText = opacity.ToString(System.Globalization.CultureInfo.InvariantCulture);

            var parser =
                WpfMath.Parsers.WpfTeXFormulaParser.Instance;

            var formula =
                parser.Parse(latex);

            double scale = 30.0;

            var environment =
                WpfMath.Rendering.WpfTeXEnvironment.Create(
                    scale: scale
                );

            var geometry =
                formula.RenderToGeometry(
                    environment,
                    scale: scale
                );

            var bounds = geometry.Bounds;

            double padding = 2.0;

            double width =
                bounds.Width + padding * 2;

            double height =
                bounds.Height + padding * 2;

            // Formel so verschieben, dass sie bei 0/0 beginnt
            double translateX =
                -bounds.X + padding;

            double translateY =
                -bounds.Y + padding;

            var converter =
                new WpfMath.Converters.SVGConverter();

            string svgPathText =
                converter.ConvertGeometry(geometry);

            string widthText =
                width.ToString(
                    System.Globalization.CultureInfo.InvariantCulture
                );

            string heightText =
                height.ToString(
                    System.Globalization.CultureInfo.InvariantCulture
                );

            string translateXText =
                translateX.ToString(
                    System.Globalization.CultureInfo.InvariantCulture
                );

            string translateYText =
                translateY.ToString(
                    System.Globalization.CultureInfo.InvariantCulture
                );

            var svgBuilder =
                new System.Text.StringBuilder();

            svgBuilder
                .AppendLine(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>"
                )
                .AppendLine(
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" " +
                    "version=\"1.1\" " +
                    "width=\"" + widthText + "\" " +
                    "height=\"" + heightText + "\" " +
                    "viewBox=\"0 0 " +
                    widthText + " " +
                    heightText + "\" " +
                    "preserveAspectRatio=\"xMidYMid meet\">"
                )
                .AppendLine(
                    "<g transform=\"translate(" +
                    translateXText + " " +
                    translateYText + ")\" " + 
                    "opacity=\"" + opacityText + "\">"
                )
                .AppendLine(svgPathText)
                .AppendLine("</g>")
                .AppendLine("</svg>");

            string filePath =
                System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "PowerPointLatexPlugin_formula.svg"
                );

            System.IO.File.WriteAllText(
                filePath,
                svgBuilder.ToString()
            );

            return filePath;
        }
    }
}