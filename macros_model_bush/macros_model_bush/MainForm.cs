using System;
using System.Windows.Forms;
using static System.Math;
//using SolidWorks.Interop.sldworks;
//using SolidWorks.Interop.swcommands;
//using SolidWorks.Interop.swconst;

namespace sw
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // forbidding anything but digits
        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox1.Text, "[^0-9]"))
            {
                textBox1.Text = textBox1.Text.Remove(textBox1.Text.Length - 1);
            }
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox2.Text, "[^0-9]"))
            {
                textBox2.Text = textBox2.Text.Remove(textBox2.Text.Length - 1);
            }
        }

        private void textBox3_TextChanged_1(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(textBox3.Text, "[^0-9]"))
            {
                textBox3.Text = textBox3.Text.Remove(textBox3.Text.Length - 1);
            }
        }

        // reference to SW window
        SldWorks swApp;

        private void button1_Click_1(object sender, EventArgs e)
        {
            // preventing empty fields
            if (string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(textBox2.Text)
                || string.IsNullOrWhiteSpace(textBox3.Text))
            {
                MessageBox.Show("Enter all values", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double x1 = Convert.ToDouble(textBox1.Text);
            double x2 = Convert.ToDouble(textBox2.Text);
            double x3 = Convert.ToDouble(textBox3.Text);

            // checking detail size condition
            if (x2 >= 42 - 5)
            {
                MessageBox.Show("Width of base can not exceed or equal the diameter of the cylinder",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x2 < 10)
            {
                MessageBox.Show("Width of base can not be less than the width of the rib",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x1 <= 42)
            {
                MessageBox.Show("Length of base can not be less than or equal " +
                    "to the diameter of the cylinder!",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (x3 <= 19)
            {
                MessageBox.Show("Height of detail can not be less than or equal " +
                    "to the length of the inner section of the cylinder!",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // m to mm conversion
            x1 = -x1 / 1000;
            x2 = -x2 / 1000;
            x3 = -x3 / 1000;

            // Opening SW 2016
            Guid Guid1 = new Guid("F16137AD-8EE8-4D2A-8CAC-DFF5D1F67522");
            object processSW = System.Activator.CreateInstance(System.Type.GetTypeFromCLSID(Guid1));

            swApp = (SldWorks)processSW;
            swApp.Visible = true;

            // selection of detail modeling mode
            swApp.NewPart();

            ModelDoc2 swDoc = null;
            bool boolstatus = false;

            swDoc = ((ModelDoc2)(swApp.ActiveDoc));

            // selecting a sketch from above, creating a circle and extruding it
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            SketchSegment skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0, 0, 0, -0.021, 0, 0)));
            swDoc.ShowNamedView2("*Trimetric", 8);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            Feature myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, Abs(x3),
                0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // selecting a sketch from above, creating a base (rectangle), adding lines of contact with a circle
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstToRectEntity)),
                ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType)),
                ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            Array vSkLines = null;
            vSkLines = ((Array)(swDoc.SketchManager.CreateCenterRectangle(0, 0, 0, x1 / 2, x2 / 2, 0)));
            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, Abs(x2 / 2), 0.000000, 0.000000, 0.021000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(0.000000, 0.021000, 0.000000, Abs(x1 / 2), Abs(x2 / 2), 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, x2 / 2, 0.000000, 0.000000, -0.021000, 0.000000)));
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(0.000000, -0.021000, 0.000000, Abs(x1 / 2), x2 / 2, 0.000000)));
            swDoc.ClearSelection2(true);

            // cutting of internal lines for extrusion
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", -0.0065662815082737427, 0, x2 / 2, false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", -0.0046846865546491534, 0, Abs(x2 / 2), false, 2, null, 0);
            boolstatus = swDoc.SketchManager.SketchTrim(4, 0, 0, 0);
            swDoc.ClearSelection2(true);

            // extruding base
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.01, Abs(x3),
                false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // selecting a sketch from above, creating an edge (rectangle) and extruding it
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstToRectEntity)),
                ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            boolstatus = swDoc.Extension.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swSketchAddConstLineDiagonalType)),
                ((int)(swUserPreferenceOption_e.swDetailingNoOptionSpecified)), true);
            vSkLines = ((Array)(swDoc.SketchManager.CreateCenterRectangle(0, 0, 0, x1 / 2, -0.005, 0)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line5", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line6", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Point1", "SKETCHPOINT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line2", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line4", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            boolstatus = swDoc.Extension.SelectByID2("Line3", "SKETCHSEGMENT", 0, 0, 0, true, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, Abs(x3), 0.01,
                false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // approximation
            swDoc.ViewZoomtofit2();
            swDoc.ViewZoomtofit2();
            swDoc.ViewZoomtofit2();

            // selection of a frontal sketch, creation of lines for diagonal cutting of an edge in height
            boolstatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(x1 / 2, 0.010000, 0.000000, -0.020396, Abs(x3), 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(false, false, false, 1, 1, +x3 - 10, +x3 - 10,
                false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // selection of frontal view, creation of the second line for diagonal cutting of the edge in height
            boolstatus = swDoc.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateLine(Abs(x1 / 2), 0.010000, 0.000000, 0.020396, Abs(x3), 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Line1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(false, true, false, 1, 1, Abs(x3) - 10, Abs(x3) - 10,
                false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // selecting top view, sketch top, creating a circle for the inner notch of the cylinder
            swDoc.ShowNamedView2("*Top", 5);
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            swDoc.ClearSelection(false)
            skSegment = null;
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.016, 0.000000, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = null;
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 1, 0, 0.01, 0.01,
                false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // filling the circle with the same radius for the cutout 17 mm below
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.016, 0.000000, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureExtrusion2(true, false, false, 0, 0, 0.017000000000000001,
                0.01, false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, true, true, true, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;

            // creating a circle of smaller radius for another cutout
            boolstatus = swDoc.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            swDoc.SketchManager.InsertSketch(true);
            swDoc.ClearSelection2(true);
            skSegment = ((SketchSegment)(swDoc.SketchManager.CreateCircle(0.000000, 0.000000, 0.000000, -0.009, 0.000000, 0.000000)));
            swDoc.ClearSelection2(true);
            boolstatus = swDoc.Extension.SelectByID2("Arc1", "SKETCHSEGMENT", 0, 0, 0, false, 0, null, 0);
            myFeature = ((Feature)(swDoc.FeatureManager.FeatureCut3(true, false, true, 1, 0, 0.017000000000000001,
                0.017000000000000001, false, false, false, false, 0.017453292519943334, 0.017453292519943334,
                false, false, false, false, false, true, true, true, true, false, 0, 0, false)));
            swDoc.ISelectionManager.EnableContourSelection = false;
        }
    }
}
