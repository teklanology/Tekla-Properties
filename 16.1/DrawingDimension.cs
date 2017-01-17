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
    class DrawingDimension
    {
        DrawingHandler drawingHandler = new DrawingHandler();

        private DimensionUnitsEnum dimensionUnits;
        private DrawingColors colour;
        private string fontName;
        private string fontHeight;
        private DrawingColors fontColour;
        private Bool fontBold;
        private Bool fontItalic;
        private PlacingEnum placing;
        private Bool placingDirectionNegative;
        private Bool placingDirectionPositive;
        private string minimalDistance;
        private string searchMargin;

        [Category("Format")]
        public DimensionUnitsEnum DimensionUnits
        {
            get { return dimensionUnits; }
            set { dimensionUnits = value; }
        }

        [Category("Appearance")]
        public DrawingColors Colour
        {
            get { return colour; }
            set { colour = value; }
        }

        [Category("Text")]
        public string FontName
        {
            get { return fontName; }
            set { fontName = value; }
        }

        [Category("Text")]
        public string FontHeight
        {
            get { return fontHeight; }
            set { fontHeight = value; }
        }

        [Category("Text")]
        public DrawingColors FontColour
        {
            get { return fontColour; }
            set { fontColour = value; }
        }

        [Category("Text")]
        public Bool FontBold
        {
            get { return fontBold; }
            set { fontBold = value; }
        }

        [Category("Text")]
        public Bool FontItalic
        {
            get { return fontItalic; }
            set { fontItalic = value; }
        }

        [Category("Placing")]
        public PlacingEnum Placing
        {
            get { return placing; }
            set { placing = value; }
        }

        [Category("Placing")]
        public Bool PlacingDirectionNegative
        {
            get { return placingDirectionNegative; }
            set { placingDirectionNegative = value; }
        }

        [Category("Placing")]
        public Bool PlacingDirectionPositive
        {
            get { return placingDirectionPositive; }
            set { placingDirectionPositive = value; }
        }

        [Category("Placing")]
        public string MinimalDistance
        {
            get { return minimalDistance; }
            set { minimalDistance = value; }
        }

        [Category("Placing")]
        public string SearchMargin
        {
            get { return searchMargin; }
            set { searchMargin = value; }
        }

        public class Angle : DrawingDimension
        {
            public void GetProperties()
            {
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                if (drawingObjectEnum.GetSize() == 1)
                {
                    while (drawingObjectEnum.MoveNext())
                    {
                        if (drawingObjectEnum.Current is Tekla.Structures.Drawing.AngleDimension)
                        {
                            Tekla.Structures.Drawing.AngleDimension angleDim = (Tekla.Structures.Drawing.AngleDimension)drawingObjectEnum.Current;

                            if (angleDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Automatic) dimensionUnits = DimensionUnitsEnum.Automatic;
                            else if (angleDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Centimeter) dimensionUnits = DimensionUnitsEnum.Centimeter;
                            else if (angleDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Inch) dimensionUnits = DimensionUnitsEnum.Inch;
                            else if (angleDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Meter) dimensionUnits = DimensionUnitsEnum.Meter;
                            else if (angleDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Millimeter) dimensionUnits = DimensionUnitsEnum.Millimeter;

                            if (angleDim.Attributes.Placing.Direction.Negative) placingDirectionNegative = Bool.True; else placingDirectionNegative = Bool.False;
                            if (angleDim.Attributes.Placing.Direction.Positive) placingDirectionPositive = Bool.True; else placingDirectionPositive = Bool.False;
                            minimalDistance = angleDim.Attributes.Placing.Distance.MinimalDistance.ToString();
                            searchMargin = angleDim.Attributes.Placing.Distance.SearchMargin.ToString();
                            if (angleDim.Attributes.Placing.Placing == DimensionSetBaseAttributes.Placings.Free) placing = PlacingEnum.Free; else placing = PlacingEnum.Fixed;

                            colour = angleDim.Attributes.Color;
                            fontName = angleDim.Attributes.Text.Font.Name;
                            fontHeight = angleDim.Attributes.Text.Font.Height.ToString();
                            fontColour = angleDim.Attributes.Text.Font.Color;
                            if (angleDim.Attributes.Text.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;
                            if (angleDim.Attributes.Text.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;

                            //distance = radialDim.Distance.ToString("F02");
                        }
                    }
                }
                if (drawingObjectEnum.GetSize() > 1)
                {
                    //distance = "";
                    dimensionUnits = new DimensionUnitsEnum();
                    colour = new DrawingColors();
                    fontName = "";
                    fontHeight = "";
                    fontColour = new DrawingColors();
                    fontItalic = new Bool();
                    fontBold = new Bool();
                    placing = new PlacingEnum();
                    placingDirectionNegative = new Bool();
                    placingDirectionPositive = new Bool();
                    minimalDistance = "";
                    searchMargin = "";
                }
            }

            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                Drawing drawing = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.AngleDimension)
                    {
                        Tekla.Structures.Drawing.AngleDimension radialDim = (Tekla.Structures.Drawing.AngleDimension)drawingObjectEnum.Current;
                        if (label == "DimensionUnits")
                        {
                            if (dimensionUnits == DimensionUnitsEnum.Automatic) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Automatic;
                            else if (dimensionUnits == DimensionUnitsEnum.Centimeter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Centimeter;
                            else if (dimensionUnits == DimensionUnitsEnum.Inch) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Inch;
                            else if (dimensionUnits == DimensionUnitsEnum.Meter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Meter;
                            else if (dimensionUnits == DimensionUnitsEnum.Millimeter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Millimeter;
                        }
                        if (label == "Colour") radialDim.Attributes.Color = colour;
                        if (label == "FontName") radialDim.Attributes.Text.Font.Name = fontName;
                        if (label == "FontHeight") radialDim.Attributes.Text.Font.Height = double.Parse(fontHeight);
                        if (label == "FontColour") radialDim.Attributes.Text.Font.Color = fontColour;
                        if (label == "FontItalic") radialDim.Attributes.Text.Font.Italic = bool.Parse(fontItalic.ToString());
                        if (label == "FontBold") radialDim.Attributes.Text.Font.Bold = bool.Parse(fontBold.ToString());
                        if (label == "Placing")
                        {
                            if (placing == PlacingEnum.Fixed) radialDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Fixed;
                            else if (placing == PlacingEnum.Free) radialDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Free;
                        }

                        if (label == "PlacingDirectionNegative") radialDim.Attributes.Placing.Direction.Negative = bool.Parse(placingDirectionNegative.ToString());
                        if (label == "PlacingDirectionPositive") radialDim.Attributes.Placing.Direction.Positive = bool.Parse(placingDirectionPositive.ToString());
                        if (label == "MinimalDistance") radialDim.Attributes.Placing.Distance.MinimalDistance = double.Parse(minimalDistance);
                        if (label == "SearchMargin") radialDim.Attributes.Placing.Distance.SearchMargin = double.Parse(searchMargin);
                        radialDim.Modify();
                        drawing.CommitChanges();
                    }
                }
            }
        }

        public class Radial : DrawingDimension
        {
            public void GetProperties()
            {
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                if (drawingObjectEnum.GetSize() == 1)
                {
                    while (drawingObjectEnum.MoveNext())
                    {
                        if (drawingObjectEnum.Current is Tekla.Structures.Drawing.CurvedDimensionSetRadial)
                        {
                            Tekla.Structures.Drawing.CurvedDimensionSetRadial radialDim = (Tekla.Structures.Drawing.CurvedDimensionSetRadial)drawingObjectEnum.Current;
                            
                            if (radialDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Automatic) dimensionUnits = DimensionUnitsEnum.Automatic;
                            else if (radialDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Centimeter) dimensionUnits = DimensionUnitsEnum.Centimeter;
                            else if (radialDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Inch) dimensionUnits = DimensionUnitsEnum.Inch;
                            else if (radialDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Meter) dimensionUnits = DimensionUnitsEnum.Meter;
                            else if (radialDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Millimeter) dimensionUnits = DimensionUnitsEnum.Millimeter;

                            if (radialDim.Attributes.Placing.Direction.Negative) placingDirectionNegative = Bool.True; else placingDirectionNegative = Bool.False;
                            if (radialDim.Attributes.Placing.Direction.Positive) placingDirectionPositive = Bool.True; else placingDirectionPositive = Bool.False;
                            minimalDistance = radialDim.Attributes.Placing.Distance.MinimalDistance.ToString();
                            searchMargin = radialDim.Attributes.Placing.Distance.SearchMargin.ToString();
                            if (radialDim.Attributes.Placing.Placing == DimensionSetBaseAttributes.Placings.Free) placing = PlacingEnum.Free; else placing = PlacingEnum.Fixed;

                            colour = radialDim.Attributes.Color;
                            fontName = radialDim.Attributes.Text.Font.Name;
                            fontHeight = radialDim.Attributes.Text.Font.Height.ToString();
                            fontColour = radialDim.Attributes.Text.Font.Color;
                            if (radialDim.Attributes.Text.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;
                            if (radialDim.Attributes.Text.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;
                        }
                    }
                }
                if (drawingObjectEnum.GetSize() > 1)
                {
                    dimensionUnits = new DimensionUnitsEnum();
                    colour = new DrawingColors();
                    fontName = "";
                    fontHeight = "";
                    fontColour = new DrawingColors();
                    fontItalic = new Bool();
                    fontBold = new Bool();
                    placing = new PlacingEnum();
                    placingDirectionNegative = new Bool();
                    placingDirectionPositive = new Bool();
                    minimalDistance = "";
                    searchMargin = "";
                }
            }

            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                Drawing drawing = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.CurvedDimensionSetRadial)
                    {
                        Tekla.Structures.Drawing.CurvedDimensionSetRadial radialDim = (Tekla.Structures.Drawing.CurvedDimensionSetRadial)drawingObjectEnum.Current;
                        if (label == "DimensionUnits")
                        {
                            if (dimensionUnits == DimensionUnitsEnum.Automatic) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Automatic;
                            else if (dimensionUnits == DimensionUnitsEnum.Centimeter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Centimeter;
                            else if (dimensionUnits == DimensionUnitsEnum.Inch) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Inch;
                            else if (dimensionUnits == DimensionUnitsEnum.Meter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Meter;
                            else if (dimensionUnits == DimensionUnitsEnum.Millimeter) radialDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Millimeter;
                        }
                        if (label == "Colour") radialDim.Attributes.Color = colour;
                        if (label == "FontName") radialDim.Attributes.Text.Font.Name = fontName;
                        if (label == "FontHeight") radialDim.Attributes.Text.Font.Height = double.Parse(fontHeight);
                        if (label == "FontColour") radialDim.Attributes.Text.Font.Color = fontColour;
                        if (label == "FontItalic") radialDim.Attributes.Text.Font.Italic = bool.Parse(fontItalic.ToString());
                        if (label == "FontBold") radialDim.Attributes.Text.Font.Bold = bool.Parse(fontBold.ToString());
                        if (label == "Placing")
                        {
                            if (placing == PlacingEnum.Fixed) radialDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Fixed;
                            else if (placing == PlacingEnum.Free) radialDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Free;
                        }

                        if (label == "PlacingDirectionNegative") radialDim.Attributes.Placing.Direction.Negative = bool.Parse(placingDirectionNegative.ToString());
                        if (label == "PlacingDirectionPositive") radialDim.Attributes.Placing.Direction.Positive = bool.Parse(placingDirectionPositive.ToString());
                        if (label == "MinimalDistance") radialDim.Attributes.Placing.Distance.MinimalDistance = double.Parse(minimalDistance);
                        if (label == "SearchMargin") radialDim.Attributes.Placing.Distance.SearchMargin = double.Parse(searchMargin);
                        radialDim.Modify();
                        drawing.CommitChanges();
                    }
                }
            }
        }

        public class Radius : DrawingDimension
        {
            private string distance;

            public string Distance
            {
                get { return distance; }
                set { distance = value; }
            }

            public void GetProperties()
            {
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                if (drawingObjectEnum.GetSize() == 1)
                {
                    while (drawingObjectEnum.MoveNext())
                    {
                        if (drawingObjectEnum.Current is Tekla.Structures.Drawing.RadiusDimension)
                        {
                            Tekla.Structures.Drawing.RadiusDimension radiusDim = (Tekla.Structures.Drawing.RadiusDimension)drawingObjectEnum.Current;
                            
                            if (radiusDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Automatic) dimensionUnits = DimensionUnitsEnum.Automatic;
                            else if (radiusDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Centimeter) dimensionUnits = DimensionUnitsEnum.Centimeter;
                            else if (radiusDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Inch) dimensionUnits = DimensionUnitsEnum.Inch;
                            else if (radiusDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Meter) dimensionUnits = DimensionUnitsEnum.Meter;
                            else if (radiusDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Millimeter) dimensionUnits = DimensionUnitsEnum.Millimeter;

                            if (radiusDim.Attributes.Placing.Direction.Negative) placingDirectionNegative = Bool.True; else placingDirectionNegative = Bool.False;
                            if (radiusDim.Attributes.Placing.Direction.Positive) placingDirectionPositive = Bool.True; else placingDirectionPositive = Bool.False;
                            minimalDistance = radiusDim.Attributes.Placing.Distance.MinimalDistance.ToString();
                            searchMargin = radiusDim.Attributes.Placing.Distance.SearchMargin.ToString();
                            if (radiusDim.Attributes.Placing.Placing == DimensionSetBaseAttributes.Placings.Free) placing = PlacingEnum.Free; else placing = PlacingEnum.Fixed;

                            colour = radiusDim.Attributes.Color;
                            fontName = radiusDim.Attributes.Text.Font.Name;
                            fontHeight = radiusDim.Attributes.Text.Font.Height.ToString();
                            fontColour = radiusDim.Attributes.Text.Font.Color;
                            if (radiusDim.Attributes.Text.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;
                            if (radiusDim.Attributes.Text.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;

                            distance = radiusDim.Distance.ToString("F02");
                        }
                    }
                }
                if (drawingObjectEnum.GetSize() > 1)
                {
                    distance = "";
                    dimensionUnits = new DimensionUnitsEnum();
                    colour = new DrawingColors();
                    fontName = "";
                    fontHeight = "";
                    fontColour = new DrawingColors();
                    fontItalic = new Bool();
                    fontBold = new Bool();
                    placing = new PlacingEnum();
                    placingDirectionNegative = new Bool();
                    placingDirectionPositive = new Bool();
                    minimalDistance = "";
                    searchMargin = "";
                }
            }

            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                Drawing drawing = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.RadiusDimension)
                    {
                        Tekla.Structures.Drawing.RadiusDimension radiusDim = (Tekla.Structures.Drawing.RadiusDimension)drawingObjectEnum.Current;
                        if (label == "Distance") radiusDim.Distance = double.Parse(distance);
                        if (label == "DimensionUnits")
                        {
                            if (dimensionUnits == DimensionUnitsEnum.Automatic) radiusDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Automatic;
                            else if (dimensionUnits == DimensionUnitsEnum.Centimeter) radiusDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Centimeter;
                            else if (dimensionUnits == DimensionUnitsEnum.Inch) radiusDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Inch;
                            else if (dimensionUnits == DimensionUnitsEnum.Meter) radiusDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Meter;
                            else if (dimensionUnits == DimensionUnitsEnum.Millimeter) radiusDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Millimeter;
                        }
                        if (label == "Colour") radiusDim.Attributes.Color = colour;
                        if (label == "FontName") radiusDim.Attributes.Text.Font.Name = fontName;
                        if (label == "FontHeight") radiusDim.Attributes.Text.Font.Height = double.Parse(fontHeight);
                        if (label == "FontColour") radiusDim.Attributes.Text.Font.Color = fontColour;
                        if (label == "FontItalic") radiusDim.Attributes.Text.Font.Italic = bool.Parse(fontItalic.ToString());
                        if (label == "FontBold") radiusDim.Attributes.Text.Font.Bold = bool.Parse(fontBold.ToString());
                        if (label == "Placing")
                        {
                            if (placing == PlacingEnum.Fixed) radiusDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Fixed;
                            else if (placing == PlacingEnum.Free) radiusDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Free;
                        }

                        if (label == "PlacingDirectionNegative") radiusDim.Attributes.Placing.Direction.Negative = bool.Parse(placingDirectionNegative.ToString());
                        if (label == "PlacingDirectionPositive") radiusDim.Attributes.Placing.Direction.Positive = bool.Parse(placingDirectionPositive.ToString());
                        if (label == "MinimalDistance") radiusDim.Attributes.Placing.Distance.MinimalDistance = double.Parse(minimalDistance);
                        if (label == "SearchMargin") radiusDim.Attributes.Placing.Distance.SearchMargin = double.Parse(searchMargin);
                        radiusDim.Modify();
                        drawing.CommitChanges();
                    }
                }
            }
        }


        public class Straight : DrawingDimension
        {
            private DimensionSetBaseAttributes.DimensionTypes dimensionType;
            private ShortDimensionEnum shortDimensionType;

            [Category("Dimension Type")]
            public DimensionSetBaseAttributes.DimensionTypes DimensionType
            {
                get { return dimensionType; }
                set { dimensionType = value; }
            }

            [Category("Placing")]
            public ShortDimensionEnum ShortDimensionType
            {
                get { return shortDimensionType; }
                set { shortDimensionType = value; }
            }

            public void GetProperties()
            {
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                if (drawingObjectEnum.GetSize() == 1)
                {
                    while (drawingObjectEnum.MoveNext())
                    {
                        if (drawingObjectEnum.Current is Tekla.Structures.Drawing.StraightDimensionSet)
                        {
                            Tekla.Structures.Drawing.StraightDimensionSet straightDim = (Tekla.Structures.Drawing.StraightDimensionSet)drawingObjectEnum.Current;
                            dimensionType = straightDim.Attributes.DimensionType;
                            if (straightDim.Attributes.ShortDimension == DimensionSetBaseAttributes.ShortDimensionTypes.Inside) shortDimensionType = ShortDimensionEnum.Inside;
                            else shortDimensionType = ShortDimensionEnum.Outside;

                            if (straightDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Automatic) dimensionUnits = DimensionUnitsEnum.Automatic;
                            else if (straightDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Centimeter) dimensionUnits = DimensionUnitsEnum.Centimeter;
                            else if (straightDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Inch) dimensionUnits = DimensionUnitsEnum.Inch;
                            else if (straightDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Meter) dimensionUnits = DimensionUnitsEnum.Meter;
                            else if (straightDim.Attributes.Format.Unit == DimensionSetBaseAttributes.DimensionValueUnits.Millimeter) dimensionUnits = DimensionUnitsEnum.Millimeter;

                            colour = straightDim.Attributes.Color;
                            fontName = straightDim.Attributes.Text.Font.Name;
                            fontHeight = straightDim.Attributes.Text.Font.Height.ToString();
                            fontColour = straightDim.Attributes.Text.Font.Color;
                            if (straightDim.Attributes.Text.Font.Italic) fontItalic = Bool.True; else fontItalic = Bool.False;
                            if (straightDim.Attributes.Text.Font.Bold) fontBold = Bool.True; else fontBold = Bool.False;
                            if (straightDim.Attributes.Placing.Placing == DimensionSetBaseAttributes.Placings.Free) placing = PlacingEnum.Free;
                            else placing = PlacingEnum.Fixed;

                            if (straightDim.Attributes.Placing.Direction.Negative) placingDirectionNegative = Bool.True; else placingDirectionNegative = Bool.False;
                            if (straightDim.Attributes.Placing.Direction.Positive) placingDirectionPositive = Bool.True; else placingDirectionPositive = Bool.False;
                            minimalDistance = straightDim.Attributes.Placing.Distance.MinimalDistance.ToString();
                            searchMargin = straightDim.Attributes.Placing.Distance.SearchMargin.ToString();
                        }
                    }
                }
                else if (drawingObjectEnum.GetSize() > 1)
                {
                    dimensionType = new DimensionSetBaseAttributes.DimensionTypes();
                    shortDimensionType = new ShortDimensionEnum();
                    dimensionUnits = new DimensionUnitsEnum();
                    colour = new DrawingColors();
                    fontName = "";
                    fontHeight = "";
                    fontColour = new DrawingColors();
                    fontItalic = new Bool();
                    fontBold = new Bool();
                    placing = new PlacingEnum();
                    placingDirectionNegative = new Bool();
                    placingDirectionPositive = new Bool();
                    minimalDistance = "";
                    searchMargin = "";
                }
            }

            public void Modify(PropertyValueChangedEventArgs e)
            {
                string label = e.ChangedItem.Label;
                Drawing drawing = drawingHandler.GetActiveDrawing();
                DrawingObjectEnumerator drawingObjectEnum = drawingHandler.GetDrawingObjectSelector().GetSelected();
                while (drawingObjectEnum.MoveNext())
                {
                    if (drawingObjectEnum.Current is Tekla.Structures.Drawing.StraightDimensionSet)
                    {
                        Tekla.Structures.Drawing.StraightDimensionSet straightDim = (Tekla.Structures.Drawing.StraightDimensionSet)drawingObjectEnum.Current;
                        if (label == "DimensionType") straightDim.Attributes.DimensionType = dimensionType;
                        if (label == "DimensionUnits")
                        {
                            if (dimensionUnits == DimensionUnitsEnum.Automatic) straightDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Automatic;
                            else if (dimensionUnits == DimensionUnitsEnum.Centimeter) straightDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Centimeter;
                            else if (dimensionUnits == DimensionUnitsEnum.Inch) straightDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Inch;
                            else if (dimensionUnits == DimensionUnitsEnum.Meter) straightDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Meter;
                            else if (dimensionUnits == DimensionUnitsEnum.Millimeter) straightDim.Attributes.Format.Unit = DimensionSetBaseAttributes.DimensionValueUnits.Millimeter;
                        }
                        if (label == "ShortDimensionType")
                        {
                            if (shortDimensionType == ShortDimensionEnum.Inside) straightDim.Attributes.ShortDimension = DimensionSetBaseAttributes.ShortDimensionTypes.Inside;
                            else if (shortDimensionType == ShortDimensionEnum.Outside) straightDim.Attributes.ShortDimension = DimensionSetBaseAttributes.ShortDimensionTypes.Outside;
                        }
                        if (label == "Placing")
                        {
                            if (placing == PlacingEnum.Fixed) straightDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Fixed;
                            else if (placing == PlacingEnum.Free) straightDim.Attributes.Placing.Placing = DimensionSetBaseAttributes.Placings.Free;
                        }
                        if (label == "SearchMargin") straightDim.Attributes.Placing.Distance.SearchMargin = double.Parse(searchMargin);
                        if (label == "MinimalDistance") straightDim.Attributes.Placing.Distance.MinimalDistance = double.Parse(minimalDistance);
                        if (label == "PlacingDirectionPositive") straightDim.Attributes.Placing.Direction.Positive = bool.Parse(placingDirectionPositive.ToString());
                        if (label == "PlacingDirectionNegative") straightDim.Attributes.Placing.Direction.Negative = bool.Parse(placingDirectionNegative.ToString());
                        if (label == "FontColour") straightDim.Attributes.Text.Font.Color = fontColour;
                        if (label == "FontHeight") straightDim.Attributes.Text.Font.Height = double.Parse(fontHeight);
                        if (label == "FontName") straightDim.Attributes.Text.Font.Name = fontName;
                        if (label == "FontBold") straightDim.Attributes.Text.Font.Bold = bool.Parse(fontBold.ToString());
                        if (label == "FontItalic") straightDim.Attributes.Text.Font.Italic = bool.Parse(fontItalic.ToString());
                        if (label == "Colour") straightDim.Attributes.Color = colour;
                        straightDim.Modify();
                        drawing.CommitChanges();
                    }
                }
            }
        }
    }
}
