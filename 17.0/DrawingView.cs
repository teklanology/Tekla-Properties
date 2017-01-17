using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Drawing;
using Tekla.Structures.Geometry3d;
using System.Drawing.Design;
using System.Resources;
using Reflection = System.Reflection;

namespace TeklaProperties
{
    class DrawingView
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private string scale;
        private string sizeXMin;
        private string sizeXMax;
        private string sizeYMin;
        private string sizeYMax;
        private string sizeDepthDown;
        private string sizeDepthUp;
        private string viewExtension;
        private Bool fixedViewPlacing;
        private Bool cutParts;
        private string minimumLength;

        [Category("View")]
        public string Scale
        {
            get { return scale; }
            set { scale = value; }
        }

        [Category("View")]
        public string SizeXMin
        {
            get { return sizeXMin; }
            set { sizeXMin = value; }
        }

        [Category("View")]
        public string SizeXMax
        {
            get { return sizeXMax; }
            set { sizeXMax = value; }
        }

        [Category("View")]
        public string SizeYMin
        {
            get { return sizeYMin; }
            set { sizeYMin = value; }
        }

        [Category("View")]
        public string SizeYMax
        {
            get { return sizeYMax; }
            set { sizeYMax = value; }
        }

        [Category("View")]
        public string SizeDepthDown
        {
            get { return sizeDepthDown; }
            set { sizeDepthDown = value; }
        }

        [Category("View")]
        public string SizeDepthUp
        {
            get { return sizeDepthUp; }
            set { sizeDepthUp = value; }
        }

        [Category("View")]
        public string ViewExtension
        {
            get { return viewExtension; }
            set { viewExtension = value; }
        }

        [Category("View")]
        public Bool FixedViewPlacing
        {
            get { return fixedViewPlacing; }
            set { fixedViewPlacing = value; }
        }

        [Category("Shortening")]
        public Bool CutParts
        {
            get { return cutParts; }
            set { cutParts = value; }
        }

        [Category("Shortening")]
        public string MinimumLength
        {
            get { return minimumLength; }
            set { minimumLength = value; }
        }

        public void GetProperties()
        {
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            if (drawingObjectEnum.GetSize() == 1)
            {
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.View)
                    {
                        Tekla.Structures.Drawing.View drawingView = (Tekla.Structures.Drawing.View)drawingObjectEnum.Current;
                        scale = drawingView.Attributes.Scale.ToString("F02");
                        sizeXMin = drawingView.RestrictionBox.MinPoint.X.ToString("F02");
                        sizeXMax = drawingView.RestrictionBox.MaxPoint.X.ToString("F02");
                        sizeYMin = drawingView.RestrictionBox.MinPoint.Y.ToString("F02");
                        sizeYMax = drawingView.RestrictionBox.MaxPoint.Y.ToString("F02");
                        sizeDepthDown = Math.Abs(drawingView.RestrictionBox.MinPoint.Z).ToString("F02");
                        sizeDepthUp = drawingView.RestrictionBox.MaxPoint.Z.ToString("F02");
                        viewExtension = drawingView.Attributes.ViewExtensionForNeighbourParts.ToString("F02");
                        if (drawingView.Attributes.FixedViewPlacing) fixedViewPlacing = Bool.True; else fixedViewPlacing = Bool.False;
                        if (drawingView.Attributes.Shortening.CutParts) cutParts = Bool.True; else cutParts = Bool.False;
                        minimumLength = drawingView.Attributes.Shortening.MinimumLength.ToString("F02");
                    }
                }
            }
            else if (drawingObjectEnum.GetSize() > 1)
            {
                scale = "";
                sizeXMin = "";
                sizeXMax = "";
                sizeYMin = "";
                sizeYMax = "";
                sizeDepthDown = "";
                sizeDepthUp = "";
                viewExtension = "";
                fixedViewPlacing = new Bool();
                cutParts = new Bool();
                minimumLength = "";
            }
        }

        public void Modify(PropertyValueChangedEventArgs e)
        {
            string label = e.ChangedItem.Label;
            Drawing drawing = drawingHandler.GetActiveDrawing();
            DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
            while (drawingObjectEnum.MoveNext())
            {
                if (drawingObjectEnum.Current is Tekla.Structures.Drawing.View)
                {
                    Tekla.Structures.Drawing.View drawingView = (Tekla.Structures.Drawing.View)drawingObjectEnum.Current;
                    if (label == "Scale") drawingView.Attributes.Scale = double.Parse(scale);
                    if (label == "SizeXMin") drawingView.RestrictionBox.MinPoint.X = double.Parse(sizeXMin);
                    if (label == "SizeXMax") drawingView.RestrictionBox.MaxPoint.X = double.Parse(sizeXMax);
                    if (label == "SizeYMin") drawingView.RestrictionBox.MinPoint.Y = double.Parse(sizeYMin);
                    if (label == "SizeYMax") drawingView.RestrictionBox.MaxPoint.Y = double.Parse(sizeYMax);
                    if (label == "SizeDepthDown") drawingView.RestrictionBox.MinPoint.Z = double.Parse(sizeDepthDown) * -1;
                    if (label == "SizeDepthUp") drawingView.RestrictionBox.MaxPoint.Z = double.Parse(sizeDepthUp);
                    if (label == "ViewExtension") drawingView.Attributes.ViewExtensionForNeighbourParts = double.Parse(viewExtension);
                    if (label == "FixedViewPlacing") drawingView.Attributes.FixedViewPlacing = bool.Parse(fixedViewPlacing.ToString());
                    if (label == "CutParts") drawingView.Attributes.Shortening.CutParts = bool.Parse(cutParts.ToString());
                    if (label == "MinimumLength") drawingView.Attributes.Shortening.MinimumLength = double.Parse(minimumLength);
                    drawingView.Modify();
                    drawing.CommitChanges();
                }
            }
        }
    }
}
