using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Threading;
using System.Drawing;

namespace NewBTASProto
{
    public partial class Cap_Var_Form : Form
    {
        public Cap_Var_Form(int Cells, CScanDataStore lastReading, string tech, string test_type, float cellLim, string d_data, double average, string workOrder, string step)
        {
            InitializeComponent();

            //test chart
            chart1.Series.Clear();

            var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "Series1",
                Color = System.Drawing.Color.Green,
                IsVisibleInLegend = false,
                IsXValueIndexed = true,
                ChartType = SeriesChartType.Column,
                BorderColor = System.Drawing.Color.DarkGray,
                BorderWidth = 1
            };
            this.chart1.Series.Add(series1);
            chart1.ChartAreas[0].AxisX.Title = "Cells";
            chart1.ChartAreas[0].AxisY.Title = "Voltage";

            for (int i = 0; i < Cells; i++)
            {
                if (GlobalVars.Pos2Neg == false)
                {
                    series1.Points.AddXY(i + 1, lastReading.orderedCells[i]);
                    // color test
                    series1.Points[i].Color = pointColorPostTest(lastReading.orderedCells[i], tech, test_type, cellLim, d_data);
                }
                else
                {
                    series1.Points.AddXY(i + 1, lastReading.orderedCells[Cells - i - 1]);
                    // color test
                    series1.Points[i].Color = pointColorPostTest(lastReading.orderedCells[Cells - i - 1], tech, test_type, cellLim, d_data);
                }
            }


            chart1.Titles.Clear();
            // work order - step - test type - date and time
            chart1.Titles.Add("Work Order:  " + workOrder + "     Step#:  " + step + "     Test Type:  " + test_type + "     (" + System.DateTime.Now.ToString() + ")" + Environment.NewLine + Environment.NewLine + "Average:  " + average.ToString("0.000"));
            chart1.Invalidate();
            chart1.ChartAreas[0].RecalculateAxesScale();

            chart1.ChartAreas[0].AxisY.Maximum = Math.Round(average + 2.5 * (double)GlobalVars.CapTestVarValue, 3);
            chart1.ChartAreas[0].AxisY.Minimum = Math.Round(average - 2.5 * (double)GlobalVars.CapTestVarValue, 3);
            chart1.ChartAreas[0].AxisY.Interval = Math.Round((double)GlobalVars.CapTestVarValue, 3);

            //average
            LineAnnotation annotation = new LineAnnotation();
            annotation.IsSizeAlwaysRelative = false;
            annotation.AxisX = chart1.ChartAreas[0].AxisX;
            annotation.AxisY = chart1.ChartAreas[0].AxisY;
            annotation.AnchorX = 0;
            annotation.AnchorY = average;
            annotation.Height = 0;
            annotation.Width = Cells + 1;
            annotation.LineWidth = 4;
            annotation.LineColor = Color.GreenYellow;
            annotation.StartCap = LineAnchorCapStyle.None;
            annotation.EndCap = LineAnchorCapStyle.None;
            chart1.Annotations.Add(annotation);

            //upper
            LineAnnotation upper = new LineAnnotation();
            upper.IsSizeAlwaysRelative = false;
            upper.AxisX = chart1.ChartAreas[0].AxisX;
            upper.AxisY = chart1.ChartAreas[0].AxisY;
            upper.AnchorX = 0;
            upper.AnchorY = average + (double)GlobalVars.CapTestVarValue;
            upper.Height = 0;
            upper.Width = Cells + 1;
            upper.LineWidth = 4;
            upper.LineColor = Color.Crimson;
            upper.StartCap = LineAnchorCapStyle.None;
            upper.EndCap = LineAnchorCapStyle.None;
            chart1.Annotations.Add(upper);

            //lower
            LineAnnotation lower = new LineAnnotation();
            lower.IsSizeAlwaysRelative = false;
            lower.AxisX = chart1.ChartAreas[0].AxisX;
            lower.AxisY = chart1.ChartAreas[0].AxisY;
            lower.AnchorX = 0;
            lower.AnchorY = average - (double)GlobalVars.CapTestVarValue;
            lower.Height = 0;
            lower.Width = Cells + 1;
            lower.LineWidth = 4;
            lower.LineColor = Color.Crimson;
            lower.StartCap = LineAnchorCapStyle.None;
            lower.EndCap = LineAnchorCapStyle.None;
            chart1.Annotations.Add(lower);

            //should we show the warning label
            for (int i = 0; i < Cells; i++)
            {
                if (Math.Abs(average - lastReading.orderedCells[i]) > (double)GlobalVars.CapTestVarValue)
                {
                    label1.Visible = true;
                    break;
                }
            }

                //play 
                System.Media.SystemSounds.Exclamation.Play();

        }

        private Color pointColorPostTest(double Value,string tech,string test_type, float cellLim, string d_data)
        {

            // test_type is the type of test we are generating the colors for

            // Three types of batteries (NiCd, SLA and NiCd ULM) and two directions (charge discharge)

            // normal vented NiCds
            double Min1 = 0;
            double Min2 = 0;
            double Min3 = 0;
            double Min4 = 0;
            double Max = 0;

            switch (tech)
            {
                case "NiCd":
                    // Discharge
                    if (test_type == "As Received" || test_type == "Capacity-1" || test_type == "Discharge" || test_type == "Custom Cap" || (test_type == "Combo: FC-6 Cap-1" && d_data == "Complete") || test_type == "Combo: FC-6  >>Cap-1<<" || test_type == "")
                    {
                        Min4 = 1;
                        Max = 1.05;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25;
                        Min2 = 1.5;
                        Min3 = 1.55;
                        Max = ((-1 == cellLim) ? 1.75 : cellLim);

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
                case "NiCd ULM":
                    // Discharge
                    if (test_type == "As Received" || test_type == "Capacity-1" || test_type == "Discharge" || test_type == "Custom Cap" || test_type == "")
                    {
                        Min4 = 1;
                        Max = 1.05;

                        if (Value > Max) { return Color.Green; }
                        else if (Value > Min4) { return Color.Orange; }
                        else { return Color.Red; }

                    }
                    // Charge
                    else
                    {
                        Min1 = 0.25;
                        Min2 = 1.5;
                        Min3 = 1.55;
                        Max = ((-1 == cellLim) ? 1.82 : cellLim);

                        if (Value > Max) { return Color.Red; }
                        else if (Value > Min3) { return Color.Green; }
                        else if (Value > Min2) { return Color.Orange; }
                        else if (Value > Min1) { return Color.Yellow; }
                        else { return Color.Red; }
                    }
            }

            // we'll return a purple if everything goes wrong
            return System.Drawing.Color.Purple;
        }

        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintDialog MyPrintDialog = new PrintDialog();
            if (MyPrintDialog.ShowDialog() == DialogResult.OK)
            {
                // do on a helper thread...
                ThreadPool.QueueUserWorkItem(s =>
                {
                    System.Drawing.Printing.PrintDocument doc = new System.Drawing.Printing.PrintDocument();
                    doc.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(doc_PrintPage);
                    doc.DefaultPageSettings.Landscape = true;
                    doc.Print();
                });
            }
        }

        private void doc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Bitmap bmp = new Bitmap(chart1.Width, chart1.Height, chart1.CreateGraphics());
            this.Invoke((MethodInvoker)delegate()
            {
                chart1.DrawToBitmap(bmp, new Rectangle(0, 0, chart1.Width, chart1.Height));
            });
            RectangleF bounds = e.PageSettings.PrintableArea;
            float factor = ((float)bounds.Height / (float)bmp.Width);
            e.Graphics.DrawImage(bmp, bounds.Left, 100, (factor * bmp.Width), (factor * bmp.Height));
        }

        private void chart1_Click(object sender, EventArgs e)
        {
            MouseEventArgs inClick = (MouseEventArgs)e;
            if (inClick.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(MousePosition);
            }
        }

        private void saveImageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog(this);
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            // Get file name.
            string name = saveFileDialog1.FileName;
            // Write to the file name selected.
            // ... You can write the text from a TextBox instead of a string literal.
            chart1.SaveImage(name, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        private void chart1_PostPaint(object sender, ChartPaintEventArgs e)
        {

        }

    }
}
