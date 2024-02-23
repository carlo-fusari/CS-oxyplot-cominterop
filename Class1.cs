using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using System.Runtime.InteropServices;

namespace OxyPlot_COM_interop
{
    [ComVisible(true)]
    [Guid("959A3E11-7B0A-4809-A6FA-FBFAC02C85C4")]
    public interface IChartGenerator
    {
        void GenerateChart(string filePath, string title, double[] xValues, double[] yValues);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ChartGenerator : IChartGenerator
    {
        public void GenerateChart(string filePath, string title, double[] xValues, double[] yValues)
        {
            var model = new PlotModel { Title = title };
            var series = new LineSeries();

            for (int i = 0; i < xValues.Length; i++)
            {
                series.Points.Add(new DataPoint(xValues[i], yValues[i]));
            }

            model.Series.Add(series);

            var pngExporter = new PngExporter { Width = 600, Height = 400 };
            pngExporter.ExportToFile(model, filePath);
        }
    }
}
