using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Inventor;

namespace Shaft
{
    public partial class Form1 : Form
    {
        private Inventor.Application _invApp;
        private bool _started = false;
        private PartDocument partDoc;
        private PartComponentDefinition partDef;
        public Form1()
        {
            InitializeComponent();
            try
            {
                _invApp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
            }
            catch (Exception ex)
            {
                try
                {
                    Type invAppType = Type.GetTypeFromProgID("Inventor.Application");
                    _invApp = (Inventor.Application)System.Activator.CreateInstance(invAppType);
                    _invApp.Visible = true;
                    _started = true;
                }
                catch (Exception ex2)
                {
                    MessageBox.Show(ex2.ToString());
                    MessageBox.Show("Unable to get or start Inventor");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            partDoc = _invApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, _invApp.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject)) as PartDocument;

           double width = Convert.ToDouble(textBox1.Text);
           double height = Convert.ToDouble(textBox2.Text);
            double circle = Convert.ToDouble(textBox3.Text) /2;
            int teeth = Convert.ToInt32(textBox4.Text);

            double widthTooth = Convert.ToDouble(textBox5.Text) / 2;
           

            // Set a reference to the component definition.        
            PartComponentDefinition oCompDef = partDoc.ComponentDefinition;
            PlanarSketch oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes[3]);

            // Set a reference to the transient geometry object.
            TransientGeometry oTransGeom = _invApp.TransientGeometry;
            SketchLines arr = oSketch.SketchLines;
            Point2d coor1 = oTransGeom.CreatePoint2d(0, 0);
            Point2d coor2 = oTransGeom.CreatePoint2d(width, 0);//525
            // SketchLine mainLine= oSketch.SketchLines.AddByTwoPoints(coor1, coor2);

            SketchLine mainLine = arr.AddByTwoPoints(coor1, coor2);
           
            Point2d coor3 = oTransGeom.CreatePoint2d(width, 0);
            Point2d coor4 = oTransGeom.CreatePoint2d(width, height);//45

            //SketchLine Line1 = oSketch.SketchLines.AddByTwoPoints(coor3, coor4);
            SketchLine Line1 = arr.AddByTwoPoints(coor3, coor4);
            
           
            Point2d coor5 = oTransGeom.CreatePoint2d(width, height);
            Point2d coor6 = oTransGeom.CreatePoint2d(width / 1.3635, height);//385  ()
            //SketchLine Line2 = oSketch.SketchLines.AddByTwoPoints(coor5, coor6);
            SketchLine Line2 = arr.AddByTwoPoints(coor5, coor6);
         

            Point2d coor7 = oTransGeom.CreatePoint2d(width / 1.3635, height);
            Point2d coor8 = oTransGeom.CreatePoint2d(width / 1.3635, height * 1.5);
            // SketchLine Line3 = oSketch.SketchLines.AddByTwoPoints(coor7, coor8);
            SketchLine Line3 = arr.AddByTwoPoints(coor7, coor8);
            

            Point2d coor9 = oTransGeom.CreatePoint2d(width / 1.3635, height * 1.375); //55
            Point2d coor10 = oTransGeom.CreatePoint2d(width / 1.94, height * 1.75);//270 70
            //SketchLine Line4 = oSketch.SketchLines.AddByTwoPoints(coor9, coor10);
            SketchLine Line4 = arr.AddByTwoPoints(coor9, coor10);
            //

            Point2d coor11 = oTransGeom.CreatePoint2d(width / 1.94, height * 1.75);
            Point2d coor12 = oTransGeom.CreatePoint2d(width / 1.94, height * 1.375);
            //SketchLine Line5 = oSketch.SketchLines.AddByTwoPoints(coor11, coor12);
            SketchLine Line5 = arr.AddByTwoPoints(coor11, coor12);
           

            Point2d coor13 = oTransGeom.CreatePoint2d(width / 1.94, height * 1.375);
            Point2d coor14 = oTransGeom.CreatePoint2d(width - (width / 1.3635), height * 1.375);//140
            //SketchLine Line6 = oSketch.SketchLines.AddByTwoPoints(coor13, coor14);
            SketchLine Line6 = arr.AddByTwoPoints(coor13, coor14);
            

            Point2d coor15 = oTransGeom.CreatePoint2d(width - (width / 1.3635), height * 1.375);
            Point2d coor16 = oTransGeom.CreatePoint2d(width - (width / 1.3635), height);
            // SketchLine Line7 = oSketch.SketchLines.AddByTwoPoints(coor15, coor16);
            SketchLine Line7 = arr.AddByTwoPoints(coor15, coor16);
           

            Point2d coor17 = oTransGeom.CreatePoint2d(width - (width / 1.3635), height);
            Point2d coor18 = oTransGeom.CreatePoint2d(0, height);//140
            // SketchLine Line8 = oSketch.SketchLines.AddByTwoPoints(coor17, coor18);
            SketchLine Line8 = arr.AddByTwoPoints(coor17, coor18);
            

            Point2d coor19 = oTransGeom.CreatePoint2d(0, height);
            Point2d coor20 = oTransGeom.CreatePoint2d(0, 0);
            //SketchLine Line9 = oSketch.SketchLines.AddByTwoPoints(coor19, coor20);
            SketchLine Line9 = arr.AddByTwoPoints(coor19, coor20);
            
            mainLine.EndSketchPoint.Merge(Line1.StartSketchPoint);
            Line1.EndSketchPoint.Merge(Line2.StartSketchPoint);
            Line2.EndSketchPoint.Merge(Line3.StartSketchPoint);
            Line3.EndSketchPoint.Merge(Line4.StartSketchPoint);
            Line4.EndSketchPoint.Merge(Line5.StartSketchPoint);
            Line5.EndSketchPoint.Merge(Line6.StartSketchPoint);
            Line6.EndSketchPoint.Merge(Line7.StartSketchPoint);
            Line7.EndSketchPoint.Merge(Line8.StartSketchPoint);
            Line8.EndSketchPoint.Merge(Line9.StartSketchPoint);
            Line9.EndSketchPoint.Merge(mainLine.StartSketchPoint);

            oSketch.GeometricConstraints.AddHorizontal((SketchEntity)mainLine);
            oSketch.GeometricConstraints.AddVertical((SketchEntity)Line1);
            oSketch.GeometricConstraints.AddHorizontal((SketchEntity)Line2);
            oSketch.GeometricConstraints.AddVertical((SketchEntity)Line3);
            //oSketch.GeometricConstraints.AddHorizontal((SketchEntity)Line4);
            oSketch.GeometricConstraints.AddVertical((SketchEntity)Line5);
           // oSketch.GeometricConstraints.AddHorizontal((SketchEntity)Line6);
           oSketch.GeometricConstraints.AddVertical((SketchEntity)Line7);
            oSketch.GeometricConstraints.AddParallel((SketchEntity)Line6, (SketchEntity)Line8);
            //oSketch.GeometricConstraints.AddHorizontal((SketchEntity)Line8);
            oSketch.GeometricConstraints.AddVertical((SketchEntity)Line9);
            oSketch.GeometricConstraints.AddGround((SketchEntity)mainLine.StartSketchPoint);

            Point2d text = oTransGeom.CreatePoint2d(550, 35);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line1.StartSketchPoint, Line1.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(500, 50);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line2.StartSketchPoint, Line2.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(440, 55);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line3.StartSketchPoint, Line3.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(320, 70);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line4.StartSketchPoint, Line4.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(240, 60);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line5.StartSketchPoint, Line5.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(200, 60);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line6.StartSketchPoint, Line6.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(70, 50);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line7.StartSketchPoint, Line7.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(50, 50);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line8.StartSketchPoint, Line8.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(-10, 30);
            oSketch.DimensionConstraints.AddTwoPointDistance(Line9.StartSketchPoint, Line9.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);

            text = oTransGeom.CreatePoint2d(250, -30);
            oSketch.DimensionConstraints.AddTwoPointDistance(mainLine.StartSketchPoint, mainLine.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text);


            Profile opr = oSketch.Profiles.AddForSolid();
            RevolveFeature top = oCompDef.Features.RevolveFeatures.AddFull(opr,mainLine,PartFeatureOperationEnum.kJoinOperation);




            //test
            PlanarSketch oSketch2 = oCompDef.Sketches.Add(top.Faces[8]);
            //SketchLines arr2 = oSketch2.SketchLines;
            SketchCircle oCircle = oSketch2.SketchCircles.AddByCenterRadius(oTransGeom.CreatePoint2d(0, 0), circle);//18

                                                                                                //-3.5      //16                                           //3.5
             SketchLine Line111 = oSketch2.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(-widthTooth, circle / 1.2), oTransGeom.CreatePoint2d(widthTooth, circle / 1.2));
             SketchLine Line222 = oSketch2.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(-widthTooth, circle / 0.9), oTransGeom.CreatePoint2d(widthTooth, circle / 0.9));//20
             SketchLine Line333 = oSketch2.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(-widthTooth, circle / 1.2), oTransGeom.CreatePoint2d(-widthTooth, circle / 0.9));
             SketchLine Line444 = oSketch2.SketchLines.AddByTwoPoints(oTransGeom.CreatePoint2d(widthTooth, circle / 1.2), oTransGeom.CreatePoint2d(widthTooth, circle / 0.9));

             Line111.EndSketchPoint.Merge(Line444.StartSketchPoint);
             Line444.EndSketchPoint.Merge(Line222.EndSketchPoint);
             Line222.StartSketchPoint.Merge(Line333.EndSketchPoint);
             Line333.StartSketchPoint.Merge(Line111.StartSketchPoint);

            oSketch2.GeometricConstraints.AddHorizontal((SketchEntity)Line111);
            oSketch2.GeometricConstraints.AddParallel((SketchEntity)Line111, (SketchEntity)Line222);
            oSketch2.GeometricConstraints.AddVertical((SketchEntity)Line333);
            oSketch2.GeometricConstraints.AddParallel((SketchEntity)Line333, (SketchEntity)Line444);
            oSketch2.GeometricConstraints.AddGround((SketchEntity)Line444);
            oSketch2.GeometricConstraints.AddGround((SketchEntity)oCircle.CenterSketchPoint);

            Point2d text2 = oTransGeom.CreatePoint2d(20, 25);
            oSketch2.DimensionConstraints.AddDiameter((SketchEntity)oCircle,text2);

            text2 = oTransGeom.CreatePoint2d(0, 35);
            oSketch2.DimensionConstraints.AddTwoPointDistance(Line222.StartSketchPoint, Line222.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text2);

            text2 = oTransGeom.CreatePoint2d(30, 20);
            oSketch2.DimensionConstraints.AddTwoPointDistance(Line444.StartSketchPoint, Line222.EndSketchPoint, DimensionOrientationEnum.kAlignedDim, text2);

            text2 = oTransGeom.CreatePoint2d(-25, 5);
            oSketch2.DimensionConstraints.AddTwoPointDistance(Line111.StartSketchPoint, oCircle.CenterSketchPoint, DimensionOrientationEnum.kAlignedDim, text2);

            



            Profile prof = oSketch2.Profiles.AddForSolid();

            ExtrudeFeature oExtrude = oCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(prof, width - (width / 1.3635), PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kCutOperation);
            _invApp.ActiveView.Fit();

           


          var col = _invApp.TransientObjects.CreateObjectCollection();
         col.Add(oExtrude);
                                                                                                                                    //1
          CircularPatternFeatureDefinition def = oCompDef.Features.CircularPatternFeatures.CreateDefinition(col, oCompDef.WorkAxes["X Axis"],true, teeth,360);

          CircularPatternFeature cirpat = oCompDef.Features.CircularPatternFeatures.AddByDefinition(def);

            //ModelParameter mp = partDoc.ComponentDefinition.Parameters.ModelParameters[partDoc.ComponentDefinition.Parameters.ModelParameters.Count];
            //mp.Expression = "360 deg";
            //cirpat.Parameters[1].
            cirpat.Parameters[1].Expression = "360 deg";
            //oSketch2.Edit();
            //oSketch2.ExitEdit();

            partDoc.Update();


        }
    }
}
