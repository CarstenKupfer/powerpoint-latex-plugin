using WpfMath;

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
    }
}